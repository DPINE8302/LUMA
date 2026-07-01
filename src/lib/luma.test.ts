import { describe, expect, it, beforeEach } from "vitest";
import {
  createFlashcardsTool,
  createDocumentTool,
  createChecklistItemTool,
  createGroupTool,
  createHighlightTool,
  createMaterialTool,
  createQuizTool,
  createTaskTool,
  deleteDocumentTool,
  deleteGroupMessageTool,
  deleteMaterialTool,
  deleteTaskTool,
  duplicateDocumentTool,
  moveTaskTool,
  postGroupMessageTool,
  reviewFlashcardTool,
  restoreDocumentVersionTool,
  shareMaterialToGroupTool,
  startFocusSessionTool,
  submitQuizAttemptTool,
  undoAction,
  updateChecklistItemTool,
  updateDocumentTool,
  updateGroupMemberRoleTool,
  updateGroupMessageTool,
  updateMaterialTool,
  updateTaskTool,
  updateTaskStatusTool,
  unshareMaterialFromGroupTool,
} from "./actions";
import { FallbackProvider, runLumaAgent } from "./ai";
import { completeOnboarding, createLocalAccount, signInLocalAccount, signOutLocalAccount } from "./auth";
import { SEED_VERSION_STORAGE_KEY, loadSnapshot, lumaDb, resetDatabase, seedDatabase } from "./db";
import { fileTypeFromName, materialFromFile } from "./files";
import { formatDue, gradeFlashcard, moveTaskByNaturalLanguage, tasksDueWithin } from "./scheduling";
import type { Flashcard, TaskItem } from "./types";

describe("LUMA local data and actions", () => {
  beforeEach(async () => {
    localStorage.clear();
    await resetDatabase();
  });

  it("loads the clean M.5 IM timetable preset", async () => {
    const snapshot = await resetDatabase();
    expect(snapshot.profile.name).toBe("Win Rattanaporn");
    expect(snapshot.profile.program).toBe("Innovative Multimedia Technology (IM)");
    expect(snapshot.subjects).toHaveLength(15);
    expect(snapshot.sessions).toHaveLength(17);
    expect(snapshot.tasks.map((task) => task.title)).toEqual(
      expect.arrayContaining(["English homework", "Additional Mathematics revision set", "Game design prototype pitch"]),
    );
    expect(snapshot.calendarEvents).toHaveLength(0);
    expect(snapshot.checklistItems).toHaveLength(0);
    expect(snapshot.reminders).toHaveLength(0);
    expect(snapshot.materials.map((material) => material.title)).toEqual(
      expect.arrayContaining(["Data Structures Notes - Search Algorithms", "IM Capstone Presentation Brief"]),
    );
    expect(snapshot.highlights).toHaveLength(0);
    expect(snapshot.flashcards).toHaveLength(0);
    expect(snapshot.flashcardReviews).toHaveLength(0);
    expect(snapshot.quizzes).toHaveLength(0);
    expect(snapshot.quizAttempts).toHaveLength(0);
    expect(snapshot.documentVersions).toHaveLength(0);
    expect(snapshot.groups.map((group) => group.name)).toContain("IM Capstone Crew");
    expect(snapshot.chats).toHaveLength(2);
    expect(snapshot.agentMemories).toHaveLength(2);
    expect(snapshot.agentConversations).toHaveLength(0);
    expect(snapshot.agentMessages).toHaveLength(0);
    expect(snapshot.auditLog).toHaveLength(0);
    expect(snapshot.actionHistory).toHaveLength(0);
    expect(snapshot.account.mode).toBe("demo");
    expect(snapshot.authSession.active).toBe(true);
    expect(snapshot.onboarding.completed).toBe(true);
    expect(snapshot.sessions.find((session) => session.id === "tue-art-9")?.mode).toBe("flipped");
    expect(snapshot.sessions.find((session) => session.id === "fri-health-9")?.mode).toBe("self-study");
    expect(snapshot.settings.localOnlyMaterials).toBe(true);
    expect(snapshot.settings.language).toBe("English");
    expect(snapshot.settings.backupFrequency).toBe("Weekly");
  });

  it("creates local accounts, persists sessions, and uses the active owner for new actions", async () => {
    const created = await createLocalAccount({
      name: "Maya Chen",
      email: "maya@school.test",
      password: "quietpass1",
      school: "LUMA Academy",
      gradeOrYear: "M.5",
      program: "Science",
    });
    expect(created.success).toBe(true);

    let snapshot = await loadSnapshot();
    expect(snapshot.account.email).toBe("maya@school.test");
    expect(snapshot.account.mode).toBe("local");
    expect(snapshot.profile.name).toBe("Maya Chen");
    expect(snapshot.onboarding.completed).toBe(false);

    await completeOnboarding(snapshot.account.id, ["profile", "privacy"]);
    snapshot = await loadSnapshot();
    expect(snapshot.onboarding.completed).toBe(true);

    const task = await createTaskTool({
      title: "Finish onboarding worksheet",
      subjectId: snapshot.subjects[0].id,
      dueAt: new Date().toISOString(),
    });
    expect(task.data?.meta?.ownerId).toBe(snapshot.profile.id);

    await signOutLocalAccount();
    await expect(loadSnapshot()).rejects.toThrow("auth_session_required");

    const signedIn = await signInLocalAccount("maya@school.test", "quietpass1");
    expect(signedIn.success).toBe(true);
    snapshot = await loadSnapshot();
    expect(snapshot.profile.name).toBe("Maya Chen");
    expect(snapshot.authSession.active).toBe(true);
  });

  it("isolates local account workspaces by active owner", async () => {
    const first = await createLocalAccount({
      name: "Maya Chen",
      email: "maya@school.test",
      password: "quietpass1",
    });
    expect(first.success).toBe(true);
    let snapshot = await loadSnapshot();
    await lumaDb.settings.put({ ...snapshot.settings, id: `settings-${snapshot.profile.id}`, theme: "Forest Mist" });
    await createTaskTool({ title: "Maya private task", subjectId: snapshot.subjects[0].id, dueAt: new Date().toISOString() });
    const mayaMaterial = await materialFromFile(new File(["Maya notes"], "maya.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await createMaterialTool(mayaMaterial);
    await createDocumentTool({ title: "Maya project", type: "Essay", outline: ["Claim"], body: "Maya draft" });

    const second = await createLocalAccount({
      name: "Noah Lee",
      email: "noah@school.test",
      password: "quietpass2",
    });
    expect(second.success).toBe(true);
    snapshot = await loadSnapshot();
    await lumaDb.settings.put({ ...snapshot.settings, id: `settings-${snapshot.profile.id}`, theme: "Deep Ocean" });
    await createTaskTool({ title: "Noah private task", subjectId: snapshot.subjects[0].id, dueAt: new Date().toISOString() });
    await createDocumentTool({ title: "Noah project", type: "Report", outline: ["Scope"], body: "Noah draft" });

    await signInLocalAccount("maya@school.test", "quietpass1");
    snapshot = await loadSnapshot();
    expect(snapshot.profile.name).toBe("Maya Chen");
    expect(snapshot.settings.theme).toBe("Forest Mist");
    expect(snapshot.tasks.map((task) => task.title)).toEqual(["Maya private task"]);
    expect(snapshot.materials.map((material) => material.title)).toEqual(["maya"]);
    expect(snapshot.documents.map((document) => document.title)).toEqual(["Maya project"]);

    await signInLocalAccount("noah@school.test", "quietpass2");
    snapshot = await loadSnapshot();
    expect(snapshot.profile.name).toBe("Noah Lee");
    expect(snapshot.settings.theme).toBe("Deep Ocean");
    expect(snapshot.tasks.map((task) => task.title)).toEqual(["Noah private task"]);
    expect(snapshot.materials).toHaveLength(0);
    expect(snapshot.documents.map((document) => document.title)).toEqual(["Noah project"]);
  });

  it("force-cleans stale local data once, then preserves new user data", async () => {
    const first = await resetDatabase();
    const task: TaskItem = {
      id: "old-demo-task",
      title: "Old demo task",
      subjectId: first.subjects[0].id,
      status: "todo",
      dueAt: new Date().toISOString(),
      priority: "medium",
      type: "homework",
      notes: "Should be removed by stale seed cleanup.",
    };
    await lumaDb.tasks.put(task);
    localStorage.removeItem(SEED_VERSION_STORAGE_KEY);

    const cleaned = await loadSnapshot();
    expect(cleaned.tasks.map((item) => item.title)).not.toContain("Old demo task");
    expect(cleaned.tasks.map((item) => item.title)).toContain("English homework");
    expect(cleaned.subjects).toHaveLength(15);

    const realTask: TaskItem = {
      ...task,
      id: "real-user-task",
      title: "Real user task",
      notes: "Should survive after the marker is current.",
    };
    await lumaDb.tasks.put(realTask);
    await seedDatabase();

    const reloaded = await loadSnapshot();
    expect(reloaded.tasks.map((item) => item.title)).toContain("Real user task");
  });

  it("persists expanded settings fields across reloads", async () => {
    const snapshot = await resetDatabase();
    await lumaDb.settings.put({
      ...snapshot.settings,
      language: "Thai",
      landingPage: "Calendar",
      diagnostics: true,
      localAiPerformance: "Fastest",
      focusSessionMinutes: 45,
    });

    const reloaded = await loadSnapshot();
    expect(reloaded.settings.language).toBe("Thai");
    expect(reloaded.settings.landingPage).toBe("Calendar");
    expect(reloaded.settings.diagnostics).toBe(true);
    expect(reloaded.settings.localAiPerformance).toBe("Fastest");
    expect(reloaded.settings.focusSessionMinutes).toBe(45);
  });

  it("hydrates newly added settings fields from defaults for older saved objects", async () => {
    await resetDatabase();
    await lumaDb.settings.put({
      id: "settings",
      theme: "Night Bloom",
      blur: 64,
      widgetStyle: "Glassmorphism",
      accentColor: "#746bff",
      motion: "balanced",
      focusSound: "Rain",
      localAiEndpoint: "http://localhost:11434",
      localAiModel: "llama3.2:latest",
      localOnlyMaterials: true,
      allowCloudSync: false,
      aiMemoryEnabled: true,
      experiencePreset: "Balanced",
      backgroundScene: "Night City",
      homeLayout: [],
    } as never);

    const reloaded = await loadSnapshot();
    expect(reloaded.settings.language).toBe("English");
    expect(reloaded.settings.groupAiAccess).toBe(true);
    expect(reloaded.settings.localAiPerformance).toBe("Balanced");
    expect(reloaded.settings.backupFrequency).toBe("Weekly");
  });

  it("finds active tasks due this week", async () => {
    const snapshot = await resetDatabase();
    const due = tasksDueWithin(
      [
        {
          id: "task-real-homework",
          title: "Real homework",
          subjectId: snapshot.subjects[0].id,
          status: "todo",
          dueAt: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
          priority: "medium",
          type: "homework",
          notes: "Assigned by teacher.",
        },
      ],
      7,
    );
    expect(due[0].status).not.toBe("done");
    expect(formatDue(due[0].dueAt)).toMatch(/Today|Tomorrow|[A-Z][a-z]{2}/);
  });

  it("reschedules a task from natural language", async () => {
    const snapshot = await resetDatabase();
    const now = new Date("2026-06-25T08:00:00.000Z");
    const task: TaskItem = {
      id: "task-english-homework",
      title: "English homework",
      subjectId: snapshot.subjects[0].id,
      status: "todo",
      dueAt: "2026-06-25T14:00:00.000Z",
      priority: "medium",
      type: "homework",
      notes: "Assigned by teacher.",
    };
    const moved = moveTaskByNaturalLanguage(task, "Move my English homework to tomorrow", now);
    expect(moved.dueAt).toBe("2026-06-26T14:00:00.000Z");
  });

  it("grades flashcards and updates mastery", async () => {
    const source: Flashcard = {
      id: "card-real",
      deck: "Generated Review",
      subjectId: "art-9",
      front: "What is this card for?",
      back: "Testing flashcard grading.",
      mastery: 20,
      dueAt: new Date().toISOString(),
    };
    const card = gradeFlashcard(source, "easy");
    expect(card.mastery).toBeGreaterThan(source.mastery);
    expect(new Date(card.dueAt).getTime()).toBeGreaterThan(Date.now());
  });

  it("routes fallback AI prompts into useful actions", async () => {
    const snapshot = await resetDatabase();
    const result = await new FallbackProvider().chat("What is due this week?", snapshot);
    expect(result.answer).toContain("English homework");
    expect(result.actions.map((action) => action.type)).toContain("open-route");
    expect(result.actions.map((action) => action.type)).toContain("start-focus");
  });

  it("navigates the app from natural language route requests", async () => {
    const snapshot = await resetDatabase();
    const result = await runLumaAgent("Open binary search from my notes", snapshot, snapshot.settings, "home");
    const routeAction = result.actions.find((action) => action.type === "open-route");

    expect(routeAction?.label).toBe("Open Learn");
    expect(routeAction?.payload.route).toBe("learn");
    expect(routeAction?.payload.sourceType).toBe("file");
    expect(routeAction?.payload.sourceId).toBe("material-binary-search");
  });

  it("uses the previous cited material for short follow-up actions", async () => {
    const snapshot = await resetDatabase();
    await runLumaAgent("Explain binary search from my notes", snapshot, snapshot.settings, "learn");
    const reloaded = await loadSnapshot();

    const result = await runLumaAgent("Make flashcards from it", reloaded, reloaded.settings, "learn");
    const flashcards = result.actions.find((action) => action.type === "create-flashcards");

    expect(flashcards?.payload.materialId).toBe("material-binary-search");
    expect(result.contextRefs?.some((ref) => ref.id === "material-binary-search")).toBe(true);
  });

  it("keeps group agent context limited to shared workspace sources", async () => {
    const snapshot = await resetDatabase();
    const result = await runLumaAgent("Create a meeting agenda from our group progress, shared files, and unfinished tasks.", snapshot, snapshot.settings, "together");
    const titles = result.contextRefs?.map((ref) => ref.title) ?? [];

    expect(titles).toContain("IM Capstone Crew");
    expect(titles).toContain("IM Capstone Presentation Brief");
    expect(titles).not.toContain("Data Structures Notes - Search Algorithms");
  });

  it("persists agent exchanges with bounded context and source references", async () => {
    const snapshot = await resetDatabase();
    const task = await createTaskTool({
      title: "English homework",
      subjectId: snapshot.subjects[0].id,
      dueAt: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
    });
    const updated = await loadSnapshot();

    const result = await runLumaAgent("What is due this week?", updated, updated.settings, "calendar");

    expect(result.contextChips).toContain("Calendar");
    expect(result.contextRefs?.some((ref) => ref.type === "task" && ref.id === task.data?.id)).toBe(true);
    const reloaded = await loadSnapshot();
    expect(reloaded.agentConversations).toHaveLength(1);
    expect(reloaded.agentMessages.map((message) => message.role)).toEqual(["user", "assistant"]);
    expect(reloaded.agentMessages.at(-1)?.contextRefs?.some((ref) => ref.title === "English homework")).toBe(true);
  });

  it("marks shared agent actions for confirmation before execution", async () => {
    const snapshot = await resetDatabase();
    const material = await materialFromFile(new File(["Graph algorithms"], "graphs.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await createMaterialTool(material);
    const group = await createGroupTool({
      name: "Algorithms group",
      subjectId: snapshot.subjects[0].id,
    });
    const reloaded = await loadSnapshot();

    const sharedAction = new FallbackProvider();
    const actions = await sharedAction.planActions("share graph notes with my group", reloaded);
    const share = actions.find((action) => action.type === "share-material") ?? {
      id: "manual-share",
      type: "share-material" as const,
      label: "Share file",
      payload: { materialId: material.id, groupId: group.data?.id ?? "" },
    };
    const result = await runLumaAgent("share graph notes with my group", reloaded, reloaded.settings, "together");
    const action = result.actions.find((item) => item.type === share.type);

    expect(action?.requiresConfirmation).toBe(true);
    expect(action?.confirmation?.impact).toBe("shared");
    expect(action?.preview?.impact).toBe("shared");
    expect(action?.preview?.sourceRefs?.map((ref) => ref.id)).toEqual(expect.arrayContaining([material.id, group.data?.id]));
    expect(action?.preview?.steps.some((step) => step.toLowerCase().includes("audit"))).toBe(true);
  });

  it("returns plan previews before personal agent tool actions", async () => {
    const snapshot = await resetDatabase();
    const created = await createTaskTool({
      title: "English citation worksheet",
      subjectId: snapshot.subjects[0].id,
      dueAt: "2026-06-25T14:00:00.000Z",
    });
    const reloaded = await loadSnapshot();

    const result = await runLumaAgent("Move English citation worksheet to tomorrow", reloaded, reloaded.settings, "calendar");
    const action = result.actions.find((item) => item.type === "reschedule-task");

    expect(action?.requiresConfirmation).toBe(false);
    expect(action?.preview?.title).toContain("English citation worksheet");
    expect(action?.preview?.summary).toContain("undo");
    expect(action?.preview?.sourceRefs?.[0]).toMatchObject({ type: "task", id: created.data?.id });
  });

  it("plans create-task actions from natural language with a preview", async () => {
    const snapshot = await resetDatabase();

    const result = await runLumaAgent("Add task finish Art 9 reflection tomorrow", snapshot, snapshot.settings, "calendar");
    const action = result.actions.find((item) => item.type === "create-task");

    expect(action?.label).toContain("finish Art 9 reflection");
    expect(action?.payload.title).toBe("finish Art 9 reflection");
    expect(action?.payload.subjectId).toBe("art-9");
    expect(action?.preview?.title).toContain("finish Art 9 reflection");
    expect(action?.preview?.steps).toContain("Create a private task");
  });

  it("creates tasks through a controlled tool with audit and undo", async () => {
    const snapshot = await resetDatabase();
    const result = await createTaskTool({
      title: "Finish Art reflection",
      subjectId: snapshot.subjects[0].id,
      dueAt: "2026-06-26T16:00:00.000Z",
      notes: "Created by test.",
    });

    expect(result.success).toBe(true);
    expect(result.data?.meta?.ownerId).toBe("profile-ava");
    expect(result.undo?.actionId).toBeTruthy();

    const created = await loadSnapshot();
    expect(created.tasks.map((task) => task.title)).toContain("Finish Art reflection");
    expect(created.auditLog.at(-1)?.action).toBe("create_task");
    expect(created.actionHistory.at(-1)?.inversePayload).toEqual({ deleteTaskId: result.data?.id });

    const undo = await undoAction(result.undo?.actionId ?? "");
    expect(undo.success).toBe(true);

    const undone = await loadSnapshot();
    expect(undone.tasks.map((task) => task.title)).not.toContain("Finish Art reflection");
    expect(undone.actionHistory.find((item) => item.id === result.undo?.actionId)?.status).toBe("undone");
  });

  it("moves tasks through a controlled tool and keeps an inverse action", async () => {
    const snapshot = await resetDatabase();
    const created = await createTaskTool({
      title: "English homework",
      subjectId: snapshot.subjects[0].id,
      dueAt: "2026-06-25T14:00:00.000Z",
    });

    const moved = await moveTaskTool({
      taskId: created.data?.id ?? "",
      naturalLanguageDate: "tomorrow",
      now: new Date("2026-06-25T08:00:00.000Z"),
    });

    expect(moved.success).toBe(true);
    expect(moved.data?.dueAt).toBe("2026-06-26T14:00:00.000Z");
    expect(moved.undo?.actionId).toBeTruthy();
    expect((await lumaDb.actionHistory.get(moved.undo?.actionId ?? ""))?.inversePayload).toEqual({
      taskId: created.data?.id,
      dueAt: "2026-06-25T14:00:00.000Z",
    });
  });

  it("updates, deletes, and restores tasks through controlled tools", async () => {
    const snapshot = await resetDatabase();
    const created = await createTaskTool({
      title: "Math worksheet",
      subjectId: snapshot.subjects[0].id,
      dueAt: "2026-06-25T14:00:00.000Z",
    });

    const updated = await updateTaskTool({
      taskId: created.data?.id ?? "",
      patch: { title: "Math worksheet revised", priority: "high", status: "in-progress" },
    });
    expect(updated.success).toBe(true);
    expect(updated.data?.title).toBe("Math worksheet revised");
    expect(updated.data?.priority).toBe("high");

    const deleted = await deleteTaskTool(created.data?.id ?? "");
    expect(deleted.success).toBe(true);
    expect(await lumaDb.tasks.get(created.data?.id ?? "")).toBeUndefined();

    const restored = await undoAction(deleted.undo?.actionId ?? "");
    expect(restored.success).toBe(true);
    expect((await lumaDb.tasks.get(created.data?.id ?? ""))?.title).toBe("Math worksheet revised");
  });

  it("creates and toggles checklist items through controlled tools", async () => {
    const snapshot = await resetDatabase();
    const created = await createTaskTool({
      title: "Build presentation",
      subjectId: snapshot.subjects[0].id,
      dueAt: "2026-06-25T14:00:00.000Z",
    });

    const item = await createChecklistItemTool({
      taskId: created.data?.id ?? "",
      title: "Collect screenshots",
    });
    expect(item.success).toBe(true);
    expect((await lumaDb.tasks.get(created.data?.id ?? ""))?.checklistItemIds).toContain(item.data?.id);

    const toggled = await updateChecklistItemTool({
      itemId: item.data?.id ?? "",
      patch: { done: true },
    });
    expect(toggled.data?.done).toBe(true);
  });

  it("starts focus sessions through a controlled tool with audit history", async () => {
    const snapshot = await resetDatabase();
    const result = await startFocusSessionTool({
      subjectId: snapshot.subjects[0].id,
      minutes: 25,
    });

    expect(result.success).toBe(true);
    expect(result.data?.meta?.syncState).toBe("local");
    const reloaded = await loadSnapshot();
    expect(reloaded.studySessions).toHaveLength(1);
    expect(reloaded.auditLog.at(-1)?.action).toBe("start_focus_session");
  });

  it("creates group workspaces and posts messages through shared tools", async () => {
    const snapshot = await resetDatabase();
    const group = await createGroupTool({
      name: "Biology project team",
      subjectId: snapshot.subjects[0].id,
    });

    expect(group.success).toBe(true);
    expect(group.data?.members[0].role).toBe("owner");

    const message = await postGroupMessageTool({
      groupId: group.data?.id ?? "",
      message: "I uploaded the worksheet notes.",
    });

    expect(message.success).toBe(true);
    const reloaded = await loadSnapshot();
    expect(reloaded.groups.map((item) => item.name)).toContain("Biology project team");
    expect(reloaded.chats.map((item) => item.message)).toContain("I uploaded the worksheet notes.");
    expect(reloaded.auditLog.map((entry) => entry.action)).toContain("post_group_message");
  });

  it("edits collaboration data through controlled tools with undo", async () => {
    const snapshot = await resetDatabase();
    const group = await createGroupTool({
      name: "Biology project team",
      subjectId: snapshot.subjects[0].id,
    });
    const message = await postGroupMessageTool({
      groupId: group.data?.id ?? "",
      message: "Initial update",
    });

    const edited = await updateGroupMessageTool({
      chatId: message.data?.id ?? "",
      message: "Edited update",
    });
    expect(edited.success).toBe(true);
    expect((await lumaDb.chats.get(message.data?.id ?? ""))?.message).toBe("Edited update");

    const deleted = await deleteGroupMessageTool(message.data?.id ?? "");
    expect(deleted.success).toBe(true);
    expect(await lumaDb.chats.get(message.data?.id ?? "")).toBeUndefined();
    expect((await undoAction(deleted.undo?.actionId ?? "")).success).toBe(true);
    expect((await lumaDb.chats.get(message.data?.id ?? ""))?.message).toBe("Edited update");

    const role = await updateGroupMemberRoleTool({
      groupId: group.data?.id ?? "",
      memberId: "profile-ava",
      role: "admin",
    });
    expect(role.success).toBe(true);
    expect((await lumaDb.groups.get(group.data?.id ?? ""))?.members[0].role).toBe("admin");

    const material = await materialFromFile(new File(["Shared notes"], "shared.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await createMaterialTool(material);
    await shareMaterialToGroupTool({ materialId: material.id, groupId: group.data?.id ?? "" });
    const unshared = await unshareMaterialFromGroupTool({ materialId: material.id, groupId: group.data?.id ?? "" });
    expect(unshared.success).toBe(true);
    expect(unshared.data?.sharedWith).not.toContain(group.data?.id);

    const groupTask = await createTaskTool({
      title: "Finish shared deck",
      subjectId: snapshot.subjects[0].id,
      dueAt: new Date().toISOString(),
      groupId: group.data?.id,
    });
    const moved = await updateTaskStatusTool({ taskId: groupTask.data?.id ?? "", status: "done" });
    expect(moved.success).toBe(true);
    expect(moved.data?.status).toBe("done");
  });

  it("shares study vault material with a group and can undo the share", async () => {
    const snapshot = await resetDatabase();
    const material = await materialFromFile(new File(["Cell membrane notes"], "biology.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await lumaDb.materials.put(material);
    const group = await createGroupTool({
      name: "Biology project team",
      subjectId: snapshot.subjects[0].id,
    });

    const shared = await shareMaterialToGroupTool({
      materialId: material.id,
      groupId: group.data?.id ?? "",
    });

    expect(shared.success).toBe(true);
    expect(shared.data?.sharedWith).toContain(group.data?.id);
    expect(shared.undo?.actionId).toBeTruthy();

    const undo = await undoAction(shared.undo?.actionId ?? "");
    expect(undo.success).toBe(true);
    expect((await lumaDb.materials.get(material.id))?.sharedWith).toEqual([]);
  });

  it("manages Study Vault materials through controlled tools with restore", async () => {
    const snapshot = await resetDatabase();
    const source = await materialFromFile(new File(["Binary search runs in logarithmic time."], "algorithms.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    const created = await createMaterialTool(source);
    expect(created.success).toBe(true);
    expect(created.data?.meta?.visibility).toBe("private");

    const updated = await updateMaterialTool({
      materialId: source.id,
      patch: { title: "Algorithms notes", folder: "Computer Science", tags: ["search", "exam"] },
    });
    expect(updated.success).toBe(true);
    expect(updated.data?.title).toBe("Algorithms notes");
    expect(updated.data?.tags).toContain("exam");

    const highlight = await createHighlightTool({
      materialId: source.id,
      text: "Binary search runs in logarithmic time.",
      note: "Important runtime",
    });
    expect(highlight.success).toBe(true);
    expect((await lumaDb.highlights.where("materialId").equals(source.id).toArray())[0]?.note).toBe("Important runtime");

    const deleted = await deleteMaterialTool(source.id);
    expect(deleted.success).toBe(true);
    expect(await lumaDb.materials.get(source.id)).toBeUndefined();
    expect(await lumaDb.highlights.where("materialId").equals(source.id).count()).toBe(0);

    const restored = await undoAction(deleted.undo?.actionId ?? "");
    expect(restored.success).toBe(true);
    expect((await lumaDb.materials.get(source.id))?.title).toBe("Algorithms notes");
    expect(await lumaDb.highlights.where("materialId").equals(source.id).count()).toBe(1);
  });

  it("creates, updates, duplicates, archives, and deletes Create projects through controlled tools", async () => {
    const created = await createDocumentTool({
      title: "Social media report",
      type: "Report",
      outline: ["Claim", "Evidence", "Conclusion"],
      body: "Draft body",
    });

    expect(created.success).toBe(true);
    expect(created.data?.status).toBe("draft");

    const updated = await updateDocumentTool({
      documentId: created.data?.id ?? "",
      patch: { body: "Revised body", status: "review" },
    });
    expect(updated.data?.body).toBe("Revised body");
    expect(updated.data?.status).toBe("review");

    const duplicated = await duplicateDocumentTool(created.data?.id ?? "");
    expect(duplicated.success).toBe(true);
    expect(duplicated.data?.title).toContain("Copy");

    const archived = await updateDocumentTool({
      documentId: created.data?.id ?? "",
      patch: { status: "archived" },
    });
    expect(archived.data?.status).toBe("archived");

    const deleted = await deleteDocumentTool(duplicated.data?.id ?? "");
    expect(deleted.success).toBe(true);
    expect(await lumaDb.documents.get(duplicated.data?.id ?? "")).toBeUndefined();

    const undo = await undoAction(deleted.undo?.actionId ?? "");
    expect(undo.success).toBe(true);
    expect((await lumaDb.documents.get(duplicated.data?.id ?? ""))?.title).toBe(duplicated.data?.title);
    expect((await loadSnapshot()).auditLog.map((entry) => entry.action)).toContain("delete_document");
  });

  it("stores and restores Create document versions through controlled tools", async () => {
    const created = await createDocumentTool({
      title: "Versioned essay",
      type: "Essay",
      outline: ["Intro", "Evidence"],
      body: "Original body",
    });
    const updated = await updateDocumentTool({
      documentId: created.data?.id ?? "",
      patch: { body: "Revised body", outline: ["Intro", "Evidence", "Conclusion"] },
      versionLabel: "Before revision",
    });
    expect(updated.success).toBe(true);
    const version = (await loadSnapshot()).documentVersions[0];
    expect(version.label).toBe("Before revision");
    expect(version.body).toBe("Original body");

    const restored = await restoreDocumentVersionTool(version.id);
    expect(restored.success).toBe(true);
    expect((await lumaDb.documents.get(created.data?.id ?? ""))?.body).toBe("Original body");

    const undo = await undoAction(restored.undo?.actionId ?? "");
    expect(undo.success).toBe(true);
    expect((await lumaDb.documents.get(created.data?.id ?? ""))?.body).toBe("Revised body");
  });

  it("generates a quiz from local material text", async () => {
    const snapshot = await resetDatabase();
    const quiz = await new FallbackProvider().generateQuiz("Physical Science 2 review note", snapshot.subjects[0].id);
    expect(quiz.questions).toHaveLength(2);
    expect(quiz.subjectId).toBe(snapshot.subjects[0].id);
  });

  it("creates flashcards and quizzes through controlled study tools with undo", async () => {
    const snapshot = await resetDatabase();
    const material = await materialFromFile(new File(["Cell membrane notes"], "biology.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await createMaterialTool(material);
    const provider = new FallbackProvider();
    const cards = await provider.generateFlashcards(material.content, material.subjectId);
    const quiz = await provider.generateQuiz(material.content, material.subjectId);

    const flashcards = await createFlashcardsTool({ materialId: material.id, flashcards: cards });
    const createdQuiz = await createQuizTool({ materialId: material.id, quiz });

    expect(flashcards.success).toBe(true);
    expect(createdQuiz.success).toBe(true);
    expect((await loadSnapshot()).auditLog.map((entry) => entry.action)).toEqual(expect.arrayContaining(["create_flashcards", "create_quiz"]));

    const undoneCards = await undoAction(flashcards.undo?.actionId ?? "");
    const undoneQuiz = await undoAction(createdQuiz.undo?.actionId ?? "");
    expect(undoneCards.success).toBe(true);
    expect(undoneQuiz.success).toBe(true);
    const reloaded = await loadSnapshot();
    expect(reloaded.flashcards).toHaveLength(0);
    expect(reloaded.quizzes).toHaveLength(0);
  });

  it("saves flashcard reviews and quiz attempts with undo", async () => {
    const snapshot = await resetDatabase();
    const material = await materialFromFile(new File(["Binary search runs in logarithmic time."], "algorithms.txt", { type: "text/plain" }), snapshot.subjects[0].id);
    await createMaterialTool(material);
    const provider = new FallbackProvider();
    const cards = await provider.generateFlashcards(material.content, material.subjectId);
    const quiz = await provider.generateQuiz(material.content, material.subjectId);
    const createdCards = await createFlashcardsTool({ materialId: material.id, flashcards: cards });
    const createdQuiz = await createQuizTool({ materialId: material.id, quiz });

    const card = createdCards.data?.[0];
    const review = await reviewFlashcardTool({ flashcardId: card?.id ?? "", quality: "easy" });
    expect(review.success).toBe(true);
    expect((await lumaDb.flashcards.get(card?.id ?? ""))?.mastery).toBeGreaterThan(card?.mastery ?? 0);
    expect((await loadSnapshot()).flashcardReviews).toHaveLength(1);

    const attempt = await submitQuizAttemptTool({
      quizId: createdQuiz.data?.id ?? "",
      answers: Object.fromEntries((createdQuiz.data?.questions ?? []).map((question) => [question.id, question.answer])),
    });
    expect(attempt.success).toBe(true);
    expect(attempt.data?.score).toBe(100);
    expect((await lumaDb.quizzes.get(createdQuiz.data?.id ?? ""))?.completedAt).toBeTruthy();

    const undoReview = await undoAction(review.undo?.actionId ?? "");
    const undoAttempt = await undoAction(attempt.undo?.actionId ?? "");
    expect(undoReview.success).toBe(true);
    expect(undoAttempt.success).toBe(true);
    const reloaded = await loadSnapshot();
    expect(reloaded.flashcardReviews).toHaveLength(0);
    expect(reloaded.quizAttempts).toHaveLength(0);
    expect(reloaded.quizzes[0]?.completedAt).toBeUndefined();
  });

  it("detects supported file types for extraction", () => {
    expect(fileTypeFromName("worksheet.pdf")).toBe("pdf");
    expect(fileTypeFromName("photo.png")).toBe("image");
    expect(fileTypeFromName("notes.docx")).toBe("doc");
    expect(fileTypeFromName("table.xlsx")).toBe("sheet");
  });
});
