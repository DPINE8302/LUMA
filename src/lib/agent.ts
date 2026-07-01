import { getCurrentOwnerId, lumaDb } from "./db";
import { formatDue, getNextClass, subjectById, tasksDueWithin } from "./scheduling";
import type { AgentAction, AgentMessage, AgentResult, LumaSnapshot, RouteId, SourceReference } from "./types";

const uid = (prefix: string) => `${prefix}-${crypto.randomUUID()}`;

export interface AgentContextBundle {
  mode: NonNullable<AgentResult["mode"]>;
  contextRefs: SourceReference[];
  contextChips: string[];
  answerPrefix?: string;
}

export function getRecentAgentMessages(
  snapshot: LumaSnapshot,
  scope: "personal" | "group" = "personal",
  groupId?: string,
  limit = 8,
): AgentMessage[] {
  const conversation = snapshot.agentConversations
    .filter((item) => item.scope === scope && (scope === "personal" || item.groupId === groupId))
    .sort((left, right) => new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime())[0];
  if (!conversation) return [];
  return snapshot.agentMessages
    .filter((message) => message.conversationId === conversation.id)
    .sort((left, right) => new Date(left.createdAt).getTime() - new Date(right.createdAt).getTime())
    .slice(-limit);
}

function words(input: string) {
  const stopWords = new Set(["can", "you", "your", "from", "with", "this", "that", "what", "where", "when", "into", "make", "show", "need", "help", "about", "notes", "explain"]);
  return input
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter((word) => word.length > 2 && !stopWords.has(word));
}

function scoreText(queryWords: string[], text: string) {
  const haystack = text.toLowerCase();
  return queryWords.reduce((score, word) => score + (haystack.includes(word) ? 1 : 0), 0);
}

function inferMode(input: string): AgentContextBundle["mode"] {
  const lower = input.toLowerCase();
  if (lower.includes("group") || lower.includes("team") || lower.includes("unfinished")) return "group";
  if (lower.includes("outline") || lower.includes("presentation") || lower.includes("essay") || lower.includes("write")) return "create";
  if (lower.includes("flashcard") || lower.includes("quiz") || lower.includes("explain") || lower.includes("summarize")) return "study";
  if (lower.includes("move") || lower.includes("create") || lower.includes("start") || lower.includes("share") || lower.includes("delete")) return "action";
  return "ask";
}

export function buildAgentContext(input: string, snapshot: LumaSnapshot, route?: RouteId): AgentContextBundle {
  const mode = inferMode(input);
  const queryWords = words(input);
  const refs: SourceReference[] = [];
  const chips: string[] = [];
  const due = tasksDueWithin(snapshot.tasks, input.toLowerCase().includes("week") ? 7 : 2).slice(0, 3);
  const nextClass = getNextClass(snapshot.sessions);
  const nextSubject = subjectById(snapshot.subjects, nextClass.session.subjectId);

  if (route) chips.push(route === "home" ? "Home" : route[0].toUpperCase() + route.slice(1));
  if (due.length > 0 && (mode === "ask" || mode === "action" || input.toLowerCase().includes("due") || input.toLowerCase().includes("week"))) {
    chips.push("Upcoming tasks");
    refs.push(
      ...due.map((task) => ({
        type: "task" as const,
        id: task.id,
        title: task.title,
        location: formatDue(task.dueAt),
      })),
    );
  }
  if (nextSubject && input.toLowerCase().includes("bring")) {
    chips.push("Next class");
    refs.push({
      type: "calendar",
      id: nextClass.session.id,
      title: nextSubject.name,
      location: `${nextClass.session.room} · ${nextClass.startsAt}`,
    });
  }

  const materialMatches = snapshot.materials
    .map((material) => ({
      material,
      score: scoreText(queryWords, `${material.title} ${material.folder} ${material.tags.join(" ")} ${material.content.slice(0, 800)}`),
    }))
    .filter(({ score }) => score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 3);

  if (mode !== "group" && materialMatches.length > 0 && (mode === "study" || input.toLowerCase().includes("find") || input.toLowerCase().includes("note"))) {
    chips.push("Study Vault");
    refs.push(
      ...materialMatches.map(({ material }) => ({
        type: "file" as const,
        id: material.id,
        title: material.title,
        location: material.folder,
      })),
    );
  }

  const group = snapshot.groups[0];
  if (group && mode === "group") {
    const sharedMaterials = snapshot.materials.filter((material) => material.sharedWith.includes(group.id));
    chips.push("Shared group only");
    refs.push({
      type: "group",
      id: group.id,
      title: group.name,
      location: `${sharedMaterials.length} shared file${sharedMaterials.length === 1 ? "" : "s"}`,
    });
    refs.push(
      ...sharedMaterials.slice(0, 2).map((material) => ({
        type: "file" as const,
        id: material.id,
        title: material.title,
        location: "Shared with group",
      })),
    );
  }

  if (snapshot.settings.aiMemoryEnabled && /prefer|preference|memory|focus style|tone|habit/.test(input.toLowerCase())) {
    const memory = snapshot.agentMemories.find((item) => item.enabled && (mode !== "group" || (item.category === "group" && item.groupId === group?.id)));
    if (memory) {
      chips.push("Memory enabled");
      refs.push({ type: "memory", id: memory.id, title: memory.category, location: "Enabled memory" });
    }
  } else {
    chips.push("Memory off");
  }

  return {
    mode,
    contextRefs: refs,
    contextChips: [...new Set(chips)].slice(0, 4),
  };
}

export function applyAgentPolicy(actions: AgentAction[], snapshot: LumaSnapshot): AgentAction[] {
  return actions.map((action) => {
    const task = action.type === "reschedule-task" ? snapshot.tasks.find((item) => item.id === action.payload.taskId) : undefined;
    const material =
      action.type === "create-flashcards" || action.type === "create-quiz" || action.type === "share-material"
        ? snapshot.materials.find((item) => item.id === action.payload.materialId)
        : undefined;
    const group = action.type === "share-material" ? snapshot.groups.find((item) => item.id === action.payload.groupId) : undefined;
    const title = typeof action.payload.title === "string" ? action.payload.title : "Generated Study Outline";
    const route = typeof action.payload.route === "string" ? action.payload.route : "home";
    const minutes = Number(action.payload.minutes ?? 25);
    const subject =
      typeof action.payload.subjectId === "string" ? snapshot.subjects.find((item) => item.id === action.payload.subjectId) : undefined;

    const preview: AgentAction["preview"] | undefined =
      action.type === "create-task"
        ? {
            title: `Create "${title}"`,
            summary: `LUMA will save this as a ${subject?.name ?? "study"} task and show it in Calendar, Home, and Today’s Brief when relevant.`,
            steps: ["Create a private task", "Set the due date and priority", "Save an audit entry", "Refresh Calendar and Home"],
            impact: "personal",
          }
        : action.type === "reschedule-task" && task
        ? {
            title: `Move "${task.title}"`,
            summary: "LUMA will update the task deadline, refresh Today’s Brief, and keep an undo action available.",
            steps: ["Find the editable task", "Calculate the new date from your request", "Update the task and audit log", "Refresh Home and Calendar"],
            impact: task.groupId ? "shared" : "personal",
            sourceRefs: [asTaskSource(task)],
          }
        : action.type === "create-flashcards" && material
          ? {
              title: `Create flashcards from "${material.title}"`,
              summary: "LUMA will use the saved material text to create a review deck in Learn.",
              steps: ["Read the Study Vault material", "Generate review cards", "Save the cards to your vault", "Open Learn with undo available"],
              impact: "personal",
              sourceRefs: [asMaterialSource(material)],
            }
          : action.type === "create-quiz" && material
            ? {
                title: `Generate quiz from "${material.title}"`,
                summary: "LUMA will create a quiz from the saved material and store it before showing results.",
                steps: ["Read the source material", "Generate questions and answers", "Save the quiz", "Open Learn with undo available"],
                impact: "personal",
                sourceRefs: [asMaterialSource(material)],
              }
            : action.type === "share-material" && material && group
              ? {
                  title: `Share "${material.title}"`,
                  summary: `This will deliberately share the file with ${group.name}. Group AI can use it after sharing.`,
                  steps: ["Check the target group", "Add the material to shared group files", "Write an audit entry", "Refresh Together"],
                  impact: "shared",
                  sourceRefs: [asMaterialSource(material), { type: "group", id: group.id, title: group.name }],
                }
              : action.type === "create-outline"
                ? {
                    title: `Create "${title}"`,
                    summary: "LUMA will create a draftable project outline in Create without editing existing documents.",
                    steps: ["Use selected Study Vault context", "Build an outline", "Create a new project", "Open Create with version history started"],
                    impact: "personal",
                  }
                : action.type === "start-focus"
                  ? {
                      title: `Start ${minutes}-minute focus`,
                      summary: "LUMA will start a focus session and add it to your analytics.",
                      steps: ["Create a focus session", "Save it locally", "Update Profile analytics", "Keep undo available"],
                      impact: "personal",
                    }
                  : action.type === "open-route"
                    ? {
                        title: `Open ${route}`,
                        summary: "LUMA will navigate without changing your data.",
                        steps: ["Open the requested workspace"],
                        impact: "personal",
                      }
                    : undefined;

    if (action.type === "share-material") {
      return {
        ...action,
        preview,
        style: "primary",
        requiresConfirmation: true,
        confirmation: {
          title: "Share material",
          description: "This will share a Study Vault material with a group workspace.",
          impact: "shared",
        },
      };
    }
    if (action.type === "reschedule-task") {
      return {
        ...action,
        preview,
        style: "secondary",
        requiresConfirmation: Boolean(task?.groupId),
        confirmation: task?.groupId
          ? {
              title: "Move group task",
              description: `This will update "${task.title}" for its shared workspace.`,
              impact: "shared",
            }
          : undefined,
      };
    }
    return {
      ...action,
      preview,
      style: action.type === "open-route" ? "secondary" : "primary",
    };
  });
}

function asTaskSource(task: LumaSnapshot["tasks"][number]): SourceReference {
  return {
    type: "task",
    id: task.id,
    title: task.title,
    location: task.dueAt,
  };
}

function asMaterialSource(material: LumaSnapshot["materials"][number]): SourceReference {
  return {
    type: "file",
    id: material.id,
    title: material.title,
    location: material.folder,
  };
}

export async function persistAgentExchange(input: string, result: AgentResult, scope: "personal" | "group" = "personal", groupId?: string) {
  const now = new Date().toISOString();
  const assistantAt = new Date(Date.now() + 1).toISOString();
  const existing = await lumaDb.agentConversations.orderBy("updatedAt").last();
  const conversation =
    existing && existing.scope === scope && existing.groupId === groupId
      ? { ...existing, updatedAt: now }
      : {
          id: uid("agent-conversation"),
          ownerId: getCurrentOwnerId(),
          scope,
          groupId,
          createdAt: now,
          updatedAt: now,
        };
  await lumaDb.transaction("rw", lumaDb.agentConversations, lumaDb.agentMessages, async () => {
    await lumaDb.agentConversations.put(conversation);
    await lumaDb.agentMessages.bulkPut([
      {
        id: uid("agent-message"),
        conversationId: conversation.id,
        role: "user",
        content: input,
        createdAt: now,
      },
      {
        id: uid("agent-message"),
        conversationId: conversation.id,
        role: "assistant",
        content: result.answer,
        createdAt: assistantAt,
        contextRefs: result.contextRefs,
        actionIds: result.actions.map((action) => action.id),
      },
    ]);
  });
}

export function enrichAgentResult(input: string, snapshot: LumaSnapshot, base: AgentResult, route?: RouteId): AgentResult {
  const context = buildAgentContext(input, snapshot, route);
  const actions = applyAgentPolicy(base.actions, snapshot);
  const mergedRefs = [...(base.contextRefs ?? []), ...context.contextRefs];
  return {
    ...base,
    actions,
    mode: context.mode,
    contextRefs: mergedRefs.filter((ref, index, refs) => refs.findIndex((candidate) => candidate.type === ref.type && candidate.id === ref.id) === index),
    contextChips: context.contextChips,
    status: base.provider === "ollama" ? "completed" : snapshot.settings.localOnlyMaterials ? "offline" : "completed",
  };
}
