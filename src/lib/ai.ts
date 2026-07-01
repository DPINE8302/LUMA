import { formatDue, getNextClass, moveTaskByNaturalLanguage, subjectById, tasksDueWithin } from "./scheduling";
import { enrichAgentResult, getRecentAgentMessages, persistAgentExchange } from "./agent";
import { generateBrowserLocalAgentResult, getBrowserLocalRuntimeState } from "./browserLocalAi";
import type {
  AgentAction,
  AgentResult,
  AiProviderStatus,
  CreateDocument,
  Flashcard,
  LumaSnapshot,
  Quiz,
  RouteId,
  SourceReference,
  TaskItem,
  UserSettings,
} from "./types";

export interface AiProvider {
  status(): Promise<AiProviderStatus>;
  chat(prompt: string, snapshot: LumaSnapshot): Promise<AgentResult>;
  summarize(text: string): Promise<string>;
  generateFlashcards(text: string, subjectId: string): Promise<Flashcard[]>;
  generateQuiz(text: string, subjectId: string): Promise<Quiz>;
  createOutline(title: string, type: CreateDocument["type"], sources: string[]): Promise<string[]>;
  planActions(input: string, snapshot: LumaSnapshot): Promise<AgentAction[]>;
}

export class OllamaProvider implements AiProvider {
  constructor(
    private endpoint: string,
    private model: string,
  ) {}

  async status(): Promise<AiProviderStatus> {
    try {
      const response = await fetch(`${this.endpoint}/api/tags`, { signal: AbortSignal.timeout(1600) });
      return response.ok ? "ready" : "offline";
    } catch {
      return "offline";
    }
  }

  async chat(prompt: string, snapshot: LumaSnapshot): Promise<AgentResult> {
    const system =
      "You are LUMA, a concise local-first agent for planning, learning, creating, and staying ahead. Prefer useful actions over long chat.";
    const context = snapshot.tasks
      .filter((task) => task.status !== "done")
      .slice(0, 6)
      .map((task) => `${task.title} due ${formatDue(task.dueAt)}`)
      .join("; ");

    const response = await fetch(`${this.endpoint}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: this.model,
        stream: false,
        messages: [
          { role: "system", content: system },
          { role: "user", content: `Context: ${context}\nRequest: ${prompt}` },
        ],
      }),
    });

    if (!response.ok) {
      throw new Error("Local model returned an error");
    }

    const data = (await response.json()) as { message?: { content?: string } };
    const fallback = new FallbackProvider();
    return {
      answer: data.message?.content?.trim() || "I can help with that.",
      actions: await fallback.planActions(prompt, snapshot),
      provider: "ollama",
    };
  }

  async summarize(text: string) {
    return this.simpleGenerate(`Summarize this study material in 4 concise bullets:\n${text}`);
  }

  async generateFlashcards(text: string, subjectId: string) {
    const fallback = new FallbackProvider();
    try {
      const summary = await this.simpleGenerate(`Create three flashcards from:\n${text}`);
      return fallback.generateFlashcards(summary, subjectId);
    } catch {
      return fallback.generateFlashcards(text, subjectId);
    }
  }

  async generateQuiz(text: string, subjectId: string) {
    const fallback = new FallbackProvider();
    try {
      const summary = await this.simpleGenerate(`Create a short quiz from:\n${text}`);
      return fallback.generateQuiz(summary, subjectId);
    } catch {
      return fallback.generateQuiz(text, subjectId);
    }
  }

  async createOutline(title: string, type: CreateDocument["type"], sources: string[]) {
    try {
      const result = await this.simpleGenerate(`Create a structured ${type} outline for "${title}" using: ${sources.join("\n")}`);
      return result
        .split("\n")
        .map((line) => line.replace(/^[-*\d.\s]+/, "").trim())
        .filter(Boolean)
        .slice(0, 8);
    } catch {
      return new FallbackProvider().createOutline(title, type, sources);
    }
  }

  async planActions(input: string, snapshot: LumaSnapshot) {
    return new FallbackProvider().planActions(input, snapshot);
  }

  private async simpleGenerate(prompt: string) {
    const response = await fetch(`${this.endpoint}/api/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: this.model, prompt, stream: false }),
    });
    if (!response.ok) {
      throw new Error("Local model returned an error");
    }
    const data = (await response.json()) as { response?: string };
    return data.response?.trim() || "";
  }
}

export class FallbackProvider implements AiProvider {
  async status(): Promise<AiProviderStatus> {
    return "ready";
  }

  async chat(prompt: string, snapshot: LumaSnapshot): Promise<AgentResult> {
    const lower = prompt.toLowerCase();
    const due = tasksDueWithin(snapshot.tasks, lower.includes("week") ? 7 : 2);
    const next = getNextClass(snapshot.sessions);
    const actions = await this.planActions(prompt, snapshot);
    const groupScope = isGroupIntent(prompt);
    const matchedMaterials = findMaterials(prompt, snapshot, { groupOnly: groupScope });
    const carriedMaterial = resolveMaterialForInput(prompt, snapshot, { groupOnly: groupScope });
    const firstGroup = snapshot.groups[0];

    if (lower.includes("study plan") || lower.includes("revision plan") || lower.includes("plan my") || lower.includes("what should i do")) {
      const planTasks = due.length ? due : snapshot.tasks.filter((task) => task.status !== "done").slice(0, 3);
      return {
        answer: buildStudyPlanAnswer(planTasks, snapshot),
        actions,
        provider: "fallback",
        contextRefs: [
          ...planTasks.slice(0, 3).map(asTaskRef),
          ...matchedMaterials.slice(0, 2).map(asMaterialRef),
        ],
      };
    }

    if (lower.includes("unfinished") || lower.includes("group progress") || lower.includes("meeting agenda") || lower.includes("summarize our group")) {
      const group = firstGroup;
      if (!group || !snapshot.settings.groupAiAccess) {
        return {
          answer: snapshot.settings.groupAiAccess
            ? "I do not see a shared workspace yet. Create a group or share a material, then I can summarize progress."
            : "Group AI access is off. Turn it on in AI Privacy before I use shared workspace context.",
          actions,
          provider: "fallback",
        };
      }
      const groupTasks = snapshot.tasks.filter((task) => group.taskIds.includes(task.id) || task.groupId === group.id);
      const unfinished = groupTasks.filter((task) => task.status !== "done");
      const sharedMaterials = snapshot.materials.filter((material) => material.sharedWith.includes(group.id) || group.materialIds.includes(material.id));
      return {
        answer: buildGroupAnswer(group.name, unfinished, sharedMaterials, group.milestones),
        actions,
        provider: "fallback",
        contextRefs: [
          { type: "group", id: group.id, title: group.name, location: `${group.progress}% progress` },
          ...unfinished.slice(0, 3).map(asTaskRef),
          ...sharedMaterials.slice(0, 2).map(asMaterialRef),
        ],
      };
    }

    if ((lower.includes("find") || lower.includes("where") || lower.includes("note") || lower.includes("explain") || lower.includes("summarize")) && (matchedMaterials.length > 0 || carriedMaterial)) {
      const material = matchedMaterials[0] ?? carriedMaterial;
      if (!material) throw new Error("Material disappeared while preparing answer.");
      const summary = await this.summarize(material.content);
      return {
        answer: lower.includes("explain")
          ? `Here is the clean version: ${summary}\n\nThe key move is to connect it to the exact source instead of guessing from memory.`
          : `I found "${material.title}" in ${material.folder}. ${summary}`,
        actions,
        provider: "fallback",
        contextRefs: [material, ...matchedMaterials.filter((item) => item.id !== material.id)].slice(0, 3).map(asMaterialRef),
      };
    }

    if (lower.includes("bring")) {
      const subject = subjectById(snapshot.subjects, next.session.subjectId);
      const bring = next.session.bring.length ? next.session.bring.join(", ") : "your usual class materials";
      return {
        answer: `For ${subject?.name ?? "your next focus block"}, bring ${bring}. Place: ${next.session.room}.`,
        actions,
        provider: "fallback",
        contextRefs: [{ type: "calendar", id: next.session.id, title: subject?.name ?? "Next class", location: `${next.session.start} · ${next.session.room}` }],
      };
    }

    if (lower.includes("due") || lower.includes("week")) {
      return {
        answer: due.length
          ? `You have ${due.length} active item${due.length === 1 ? "" : "s"} coming up: ${due
              .slice(0, 4)
              .map((task) => `${task.title} (${formatDue(task.dueAt)})`)
              .join(", ")}.\n\nBest move: start with ${due[0].title}, then let me block a focus session or move the task if the deadline is wrong.`
          : "No urgent work is due in that window.",
        actions,
        provider: "fallback",
        contextRefs: due.slice(0, 4).map(asTaskRef),
      };
    }

    if (lower.includes("flashcard")) {
      return {
        answer:
          carriedMaterial
            ? `I can turn "${carriedMaterial.title}" into a review deck with front/back cards and schedule it for today.`
            : "I can turn the current material or highlights into a review deck and schedule it for today.",
        actions,
        provider: "fallback",
        contextRefs: carriedMaterial ? [asMaterialRef(carriedMaterial)] : matchedMaterials.slice(0, 2).map(asMaterialRef),
      };
    }

    if (lower.includes("quiz")) {
      return {
        answer:
          carriedMaterial
            ? `I can generate a focused quiz from "${carriedMaterial.title}", then turn missed topics into flashcards.`
            : "I can generate a focused quiz from your latest Study Vault material, then turn missed topics into flashcards.",
        actions,
        provider: "fallback",
        contextRefs: carriedMaterial ? [asMaterialRef(carriedMaterial)] : matchedMaterials.slice(0, 2).map(asMaterialRef),
      };
    }

    return {
      answer:
        "I found your timetable, deadlines, study materials, and privacy settings. Ask what is due this week, make a study plan, explain a note, or summarize group progress.",
      actions,
      provider: "fallback",
    };
  }

  async summarize(text: string) {
    const sentences = text
      .replace(/\s+/g, " ")
      .split(/(?<=[.!?])\s+/)
      .filter(Boolean);
    return sentences.slice(0, 4).join(" ") || "No readable text found yet.";
  }

  async generateFlashcards(text: string, subjectId: string): Promise<Flashcard[]> {
    const clean = text.replace(/\s+/g, " ").trim();
    const topic = clean.split(/[.:\n]/)[0]?.slice(0, 64) || "Key concept";
    return [
      {
        id: crypto.randomUUID(),
        deck: "Generated Review",
        subjectId,
        front: `What is the main idea of ${topic}?`,
        back: clean.slice(0, 220) || "Review the source material and restate the central idea.",
        mastery: 25,
        dueAt: new Date().toISOString(),
      },
      {
        id: crypto.randomUUID(),
        deck: "Generated Review",
        subjectId,
        front: "Which detail should you remember for the exam?",
        back: clean.split(".").slice(1, 3).join(". ").trim() || "Identify the definition, process, or comparison from the source.",
        mastery: 20,
        dueAt: new Date().toISOString(),
      },
    ];
  }

  async generateQuiz(text: string, subjectId: string): Promise<Quiz> {
    const topic = text.split(/[.:\n]/)[0]?.slice(0, 60) || "Current material";
    return {
      id: crypto.randomUUID(),
      title: `${topic} Quiz`,
      subjectId,
      weakTopics: [],
      questions: [
        {
          id: crypto.randomUUID(),
          prompt: `What best describes ${topic}?`,
          answer: "The key concept from the selected material",
          options: ["The key concept from the selected material", "An unrelated example", "A schedule change", "A calendar filter"],
          type: "multiple-choice",
          topic,
        },
        {
          id: crypto.randomUUID(),
          prompt: "Write one sentence explaining why this matters.",
          answer: "It helps connect the material to the exam goal.",
          type: "short-answer",
          topic,
        },
      ],
    };
  }

  async createOutline(title: string, type: CreateDocument["type"], sources: string[]) {
    const sourceAnchor = sources[0]?.slice(0, 80) || "your imported notes";
    if (type === "Presentation") {
      return ["Opening claim", "Context", "Evidence slide", "Counterpoint", "Visual example", "Conclusion and next steps"];
    }
    if (type === "Project Plan") {
      return ["Goal", "Scope", "Milestones", "Roles", "Resources", "Risks", "Final delivery"];
    }
    return ["Introduction", `Background for ${title}`, `Evidence from ${sourceAnchor}`, "Key point one", "Key point two", "Counterpoint", "Conclusion"];
  }

  async planActions(input: string, snapshot: LumaSnapshot): Promise<AgentAction[]> {
    const lower = input.toLowerCase();
    const actions: AgentAction[] = [];
    const groupScope = isGroupIntent(input);
    const materialGroupOnly = groupScope && !/\bshare\b/.test(lower);
    const route = inferRouteForInput(input);
    const firstTask = findTaskForInput(input, snapshot) ?? snapshot.tasks.find((task) => task.status !== "done") ?? snapshot.tasks[0];
    const firstMaterial = resolveMaterialForInput(input, snapshot, { groupOnly: materialGroupOnly });
    const firstGroup = snapshot.groups[0];
    const taskDraft = parseTaskDraft(input, snapshot);

    if (route) {
      const focusedPayload: AgentAction["payload"] = { route };
      if (route === "learn" && firstMaterial) {
        focusedPayload.sourceType = "file";
        focusedPayload.sourceId = firstMaterial.id;
      }
      if (route === "calendar" && firstTask) {
        focusedPayload.sourceType = "task";
        focusedPayload.sourceId = firstTask.id;
      }
      actions.push({
        id: crypto.randomUUID(),
        type: "open-route",
        label: `Open ${routeLabel(route)}`,
        payload: focusedPayload,
      });
    }

    if (taskDraft && !lower.includes("move") && !lower.includes("reschedule")) {
      actions.push({
        id: crypto.randomUUID(),
        type: "create-task",
        label: `Create task: ${taskDraft.title}`,
        payload: taskDraft,
      });
    }
    if (firstTask && (lower.includes("move") || lower.includes("reschedule"))) {
      actions.push({
        id: crypto.randomUUID(),
        type: "reschedule-task",
        label: `Move ${firstTask.title}`,
        payload: { taskId: firstTask.id, input },
      });
    }
    if (firstMaterial && (lower.includes("flashcard") || lower.includes("worksheet") || lower.includes("review"))) {
      actions.push({
        id: crypto.randomUUID(),
        type: "create-flashcards",
        label: "Make flashcards",
        payload: { materialId: firstMaterial.id },
      });
    }
    if (firstMaterial && (lower.includes("quiz") || lower.includes("test me"))) {
      actions.push({
        id: crypto.randomUUID(),
        type: "create-quiz",
        label: "Generate quiz",
        payload: { materialId: firstMaterial.id },
      });
    }
    if (firstMaterial && firstGroup && /\bshare\b/.test(lower) && !lower.includes("shared file")) {
      actions.push({
        id: crypto.randomUUID(),
        type: "share-material",
        label: `Share ${firstMaterial.title}`,
        payload: { materialId: firstMaterial.id, groupId: firstGroup.id },
      });
    }
    if (lower.includes("outline") || lower.includes("presentation") || lower.includes("essay")) {
      actions.push({
        id: crypto.randomUUID(),
        type: "create-outline",
        label: lower.includes("presentation") ? "Build presentation outline" : "Create outline",
        payload: { title: lower.includes("presentation") ? "Presentation Outline" : "Generated Study Outline", type: lower.includes("presentation") ? "Presentation" : "Essay" },
      });
    }
    if (lower.includes("focus") || actions.length === 0) {
      actions.push({
        id: crypto.randomUUID(),
        type: "start-focus",
        label: "Start 25 min focus",
        payload: { minutes: 25 },
      });
    }
    if ((lower.includes("due") || lower.includes("week")) && !actions.some((action) => action.type === "open-route" && action.payload.route === "calendar")) {
      actions.push({
        id: crypto.randomUUID(),
        type: "open-route",
        label: "Open Calendar",
        payload: { route: "calendar" },
      });
    }
    return actions.slice(0, 4);
  }
}

function findTaskForInput(input: string, snapshot: LumaSnapshot) {
  const terms = input
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((term) => term.length > 2);
  return snapshot.tasks
    .map((task) => ({
      task,
      score: terms.reduce(
        (total, term) => total + (`${task.title} ${task.notes} ${task.assignee ?? ""}`.toLowerCase().includes(term) ? 1 : 0),
        0,
      ),
    }))
    .filter(({ score }) => score > 0)
    .sort((left, right) => right.score - left.score)
    .map(({ task }) => task)[0];
}

function findMaterials(input: string, snapshot: LumaSnapshot, options: { groupOnly?: boolean } = {}) {
  const stopWords = new Set(["can", "you", "your", "from", "with", "this", "that", "what", "where", "when", "into", "make", "show", "need", "help", "about", "notes", "explain"]);
  const terms = input
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((term) => term.length > 2 && !stopWords.has(term));
  const firstGroup = snapshot.groups[0];
  const materials = options.groupOnly && firstGroup
    ? snapshot.materials.filter((material) => material.sharedWith.includes(firstGroup.id) || firstGroup.materialIds.includes(material.id))
    : snapshot.materials;
  return materials
    .map((material) => ({
      material,
      score: terms.reduce(
        (total, term) =>
          total +
          (`${material.title} ${material.folder} ${material.tags.join(" ")} ${material.content}`.toLowerCase().includes(term) ? 1 : 0),
        0,
      ),
    }))
    .filter(({ score }) => score > 0)
    .sort((left, right) => right.score - left.score)
    .map(({ material }) => material);
}

function resolveMaterialForInput(input: string, snapshot: LumaSnapshot, options: { groupOnly?: boolean } = {}) {
  const direct = findMaterials(input, snapshot, options)[0];
  if (direct) return direct;
  if (!isFollowUpReference(input)) return undefined;

  const firstGroup = snapshot.groups[0];
  const scope = options.groupOnly ? "group" : "personal";
  const recentMessages = getRecentAgentMessages(snapshot, scope, firstGroup?.id, 10);
  const latestFileRef = [...recentMessages]
    .reverse()
    .flatMap((message) => message.contextRefs ?? [])
    .find((ref) => ref.type === "file");
  if (!latestFileRef) return undefined;
  const material = snapshot.materials.find((item) => item.id === latestFileRef.id);
  if (!material) return undefined;
  if (options.groupOnly && firstGroup && !material.sharedWith.includes(firstGroup.id) && !firstGroup.materialIds.includes(material.id)) {
    return undefined;
  }
  return material;
}

function isFollowUpReference(input: string) {
  return /\b(it|that|this|those|these|same|above|previous|last one|current)\b/i.test(input);
}

function isGroupIntent(input: string) {
  return /\b(group|team|together|shared|our|we|unfinished|meeting agenda|progress)\b/i.test(input);
}

function inferRouteForInput(input: string): RouteId | undefined {
  const lower = input.toLowerCase();
  const hasOpenIntent = /\b(open|go to|navigate|show|take me|switch to|bring me to)\b/.test(lower);
  if (!hasOpenIntent) return undefined;
  if (/\b(calendar|deadline|deadlines|schedule|due|class|classes)\b/.test(lower)) return "calendar";
  if (/\b(learn|study vault|vault|flashcards?|quiz|quizzes|notes?|materials?)\b/.test(lower)) return "learn";
  if (/\b(together|group|team|shared|crew)\b/.test(lower)) return "together";
  if (/\b(create|document|essay|presentation|outline|project)\b/.test(lower)) return "create";
  if (/\b(profile|analytics|streak|progress)\b/.test(lower)) return "profile";
  if (/\b(settings|privacy|local model|ai privacy|model)\b/.test(lower)) return "settings";
  if (/\b(home|dashboard|main)\b/.test(lower)) return "home";
  return undefined;
}

function routeLabel(route: RouteId) {
  return route === "home" ? "Home" : route[0].toUpperCase() + route.slice(1);
}

function asTaskRef(task: TaskItem): SourceReference {
  return { type: "task", id: task.id, title: task.title, location: formatDue(task.dueAt) };
}

function asMaterialRef(material: LumaSnapshot["materials"][number]): SourceReference {
  return { type: "file", id: material.id, title: material.title, location: material.folder };
}

function buildStudyPlanAnswer(tasks: TaskItem[], snapshot: LumaSnapshot) {
  if (tasks.length === 0) {
    return "I do not see active deadlines yet. Add one task or import a worksheet, and I can turn it into a timed study plan.";
  }
  const lines = tasks.slice(0, 3).map((task, index) => {
    const subject = subjectById(snapshot.subjects, task.subjectId);
    const minutes = task.estimatedMinutes ?? (task.priority === "high" ? 50 : 25);
    return `${index + 1}. ${minutes} min - ${task.title} (${subject?.name ?? "Study"}, ${formatDue(task.dueAt)})`;
  });
  return `Here is the tight plan I would run next:\n${lines.join("\n")}\n\nStart with the first high-priority item, keep each block small, and use the action below if you want me to start a focus session or create a task.`;
}

function buildGroupAnswer(
  name: string,
  unfinished: TaskItem[],
  materials: LumaSnapshot["materials"],
  milestones: LumaSnapshot["groups"][number]["milestones"],
) {
  const nextMilestone = milestones.find((milestone) => !milestone.done);
  const taskLine = unfinished.length
    ? unfinished.map((task) => `${task.title}${task.assignee ? ` (${task.assignee})` : ""}`).join(", ")
    : "No unfinished group tasks are visible.";
  const sourceLine = materials.length ? `I used ${materials.length} shared file${materials.length === 1 ? "" : "s"}.` : "No shared files are available yet.";
  const milestoneLine = nextMilestone ? `Next checkpoint: ${nextMilestone.label} (${formatDue(nextMilestone.dueAt)}).` : "All milestones are marked done.";
  return `${name} status: ${taskLine}\n${milestoneLine}\n${sourceLine}\n\nMeeting agenda: demo first, unblock the unfinished task, then assign one owner for the next checkpoint.`;
}

function parseTaskDraft(input: string, snapshot: LumaSnapshot): AgentAction["payload"] | null {
  const lower = input.toLowerCase();
  const isTaskIntent =
    lower.includes("create task") ||
    lower.includes("add task") ||
    lower.includes("new task") ||
    lower.includes("remind me") ||
    lower.includes("homework") ||
    lower.includes("assignment");
  if (!isTaskIntent) return null;

  const cleaned = input
    .replace(/^(please\s+)?(create|add|make|new)\s+(a\s+)?(task|homework|assignment)\s*(to|for|called|:)?\s*/i, "")
    .replace(/^(please\s+)?remind me to\s*/i, "")
    .replace(/\s+(by|before|due)\s+(today|tomorrow|tonight|this week).*$/i, "")
    .replace(/\s+(today|tomorrow|tonight|this week)$/i, "")
    .replace(/[.!?]+$/g, "")
    .trim();
  const title = cleaned || "New study task";
  const subject = findSubjectForInput(input, snapshot) ?? snapshot.subjects[0];
  const dueAt = inferDueAt(input);
  const priority = lower.includes("urgent") || lower.includes("important") ? "high" : lower.includes("low priority") ? "low" : "medium";

  return {
    title: title.slice(0, 96),
    subjectId: subject.id,
    dueAt,
    priority,
    type: lower.includes("exam") || lower.includes("test") ? "exam" : lower.includes("project") ? "project" : "homework",
  };
}

function findSubjectForInput(input: string, snapshot: LumaSnapshot) {
  const lower = input.toLowerCase();
  return snapshot.subjects.find((subject) => lower.includes(subject.name.toLowerCase()));
}

function inferDueAt(input: string) {
  const lower = input.toLowerCase();
  const due = new Date();
  if (lower.includes("tomorrow")) {
    due.setDate(due.getDate() + 1);
    due.setHours(17, 0, 0, 0);
    return due.toISOString();
  }
  if (lower.includes("tonight") || lower.includes("today")) {
    due.setHours(20, 0, 0, 0);
    return due.toISOString();
  }
  if (lower.includes("this week")) {
    due.setDate(due.getDate() + 3);
    due.setHours(17, 0, 0, 0);
    return due.toISOString();
  }
  due.setDate(due.getDate() + 1);
  due.setHours(17, 0, 0, 0);
  return due.toISOString();
}

export async function runLumaAgent(input: string, snapshot: LumaSnapshot, settings: UserSettings, route?: RouteId): Promise<AgentResult> {
  const scope = route === "together" || isGroupIntent(input) ? "group" : "personal";
  const groupId = snapshot.groups[0]?.id;
  const recentMessages = getRecentAgentMessages(snapshot, scope, groupId, 8);
  if (settings.localAiEndpoint === "browser-webllm" && getBrowserLocalRuntimeState().ready) {
    try {
      const result = await generateBrowserLocalAgentResult(input, snapshot, settings.localAiModel, { scope, recentMessages });
      result.actions = await new FallbackProvider().planActions(input, snapshot);
      const enriched = enrichAgentResult(input, snapshot, result, route);
      await persistAgentExchange(input, enriched, enriched.mode === "group" ? "group" : "personal", groupId);
      return enriched;
    } catch {
      const result = await new FallbackProvider().chat(input, snapshot);
      const enriched = enrichAgentResult(input, snapshot, { ...result, status: "offline" }, route);
      await persistAgentExchange(input, enriched, enriched.mode === "group" ? "group" : "personal", groupId);
      return enriched;
    }
  }
  const local = new OllamaProvider(settings.localAiEndpoint, settings.localAiModel);
  let result: AgentResult;
  if ((await local.status()) === "ready") {
    try {
      result = await local.chat(input, snapshot);
      const enriched = enrichAgentResult(input, snapshot, result, route);
      await persistAgentExchange(input, enriched, enriched.mode === "group" ? "group" : "personal", groupId);
      return enriched;
    } catch {
      result = await new FallbackProvider().chat(input, snapshot);
      const enriched = enrichAgentResult(input, snapshot, result, route);
      await persistAgentExchange(input, enriched, enriched.mode === "group" ? "group" : "personal", groupId);
      return enriched;
    }
  }
  result = await new FallbackProvider().chat(input, snapshot);
  const enriched = enrichAgentResult(input, snapshot, result, route);
  await persistAgentExchange(input, enriched, enriched.mode === "group" ? "group" : "personal", groupId);
  return enriched;
}

export function applyAgentAction(task: TaskItem, input: string) {
  return moveTaskByNaturalLanguage(task, input);
}
