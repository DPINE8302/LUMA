import Dexie, { type Table } from "dexie";
import { TIMETABLE_SEED_VERSION, createSeedData } from "../data/seed";
import type {
  ActionHistoryEntry,
  AgentConversation,
  AgentMemory,
  AgentMessage,
  AuthSession,
  AuditLogEntry,
  CalendarEvent,
  ChatMessage,
  ChecklistItem,
  ClassSession,
  CreateDocument,
  DocumentVersion,
  Flashcard,
  FlashcardReview,
  GroupWorkspace,
  Highlight,
  LocalAccount,
  LumaSnapshot,
  Material,
  OnboardingState,
  Quiz,
  QuizAttempt,
  Reminder,
  StudySession,
  Subject,
  TaskItem,
  UserProfile,
  UserSettings,
} from "./types";

export class LumaDatabase extends Dexie {
  accounts!: Table<LocalAccount, string>;
  authSessions!: Table<AuthSession, string>;
  onboarding!: Table<OnboardingState, string>;
  profile!: Table<UserProfile, string>;
  subjects!: Table<Subject, string>;
  sessions!: Table<ClassSession, string>;
  calendarEvents!: Table<CalendarEvent, string>;
  tasks!: Table<TaskItem, string>;
  checklistItems!: Table<ChecklistItem, string>;
  reminders!: Table<Reminder, string>;
  materials!: Table<Material, string>;
  highlights!: Table<Highlight, string>;
  flashcards!: Table<Flashcard, string>;
  flashcardReviews!: Table<FlashcardReview, string>;
  quizzes!: Table<Quiz, string>;
  quizAttempts!: Table<QuizAttempt, string>;
  documents!: Table<CreateDocument, string>;
  documentVersions!: Table<DocumentVersion, string>;
  groups!: Table<GroupWorkspace, string>;
  chats!: Table<ChatMessage, string>;
  studySessions!: Table<StudySession, string>;
  agentMemories!: Table<AgentMemory, string>;
  agentConversations!: Table<AgentConversation, string>;
  agentMessages!: Table<AgentMessage, string>;
  auditLog!: Table<AuditLogEntry, string>;
  actionHistory!: Table<ActionHistoryEntry, string>;
  settings!: Table<UserSettings, string>;

  constructor() {
    super("luma");
    this.version(1).stores({
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      tasks: "id, subjectId, status, dueAt, groupId",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      quizzes: "id, subjectId, completedAt",
      documents: "id, type, status, dueAt",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      settings: "id",
    });
    this.version(2).stores({
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      calendarEvents: "id, ownerId, startAt, subjectId, taskId, groupId, syncState",
      tasks: "id, subjectId, status, dueAt, groupId",
      checklistItems: "id, ownerId, taskId, done, syncState",
      reminders: "id, ownerId, remindAt, targetType, targetId, enabled, syncState",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      quizzes: "id, subjectId, completedAt",
      documents: "id, type, status, dueAt",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      agentMemories: "id, ownerId, category, groupId, enabled, updatedAt",
      auditLog: "id, ownerId, action, impact, createdAt, affectedType, affectedId",
      actionHistory: "id, ownerId, action, status, createdAt, expiresAt",
      settings: "id",
    });
    this.version(3).stores({
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      calendarEvents: "id, ownerId, startAt, subjectId, taskId, groupId, syncState",
      tasks: "id, subjectId, status, dueAt, groupId",
      checklistItems: "id, ownerId, taskId, done, syncState",
      reminders: "id, ownerId, remindAt, targetType, targetId, enabled, syncState",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      quizzes: "id, subjectId, completedAt",
      documents: "id, type, status, dueAt",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      agentMemories: "id, ownerId, category, groupId, enabled, updatedAt",
      agentConversations: "id, ownerId, scope, groupId, updatedAt",
      agentMessages: "id, conversationId, role, createdAt",
      auditLog: "id, ownerId, action, impact, createdAt, affectedType, affectedId",
      actionHistory: "id, ownerId, action, status, createdAt, expiresAt",
      settings: "id",
    });
    this.version(4).stores({
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      calendarEvents: "id, ownerId, startAt, subjectId, taskId, groupId, syncState",
      tasks: "id, subjectId, status, dueAt, groupId",
      checklistItems: "id, ownerId, taskId, done, syncState",
      reminders: "id, ownerId, remindAt, targetType, targetId, enabled, syncState",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      flashcardReviews: "id, ownerId, flashcardId, subjectId, reviewedAt, syncState",
      quizzes: "id, subjectId, completedAt",
      quizAttempts: "id, ownerId, quizId, subjectId, submittedAt, score, syncState",
      documents: "id, type, status, dueAt",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      agentMemories: "id, ownerId, category, groupId, enabled, updatedAt",
      agentConversations: "id, ownerId, scope, groupId, updatedAt",
      agentMessages: "id, conversationId, role, createdAt",
      auditLog: "id, ownerId, action, impact, createdAt, affectedType, affectedId",
      actionHistory: "id, ownerId, action, status, createdAt, expiresAt",
      settings: "id",
    });
    this.version(5).stores({
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      calendarEvents: "id, ownerId, startAt, subjectId, taskId, groupId, syncState",
      tasks: "id, subjectId, status, dueAt, groupId",
      checklistItems: "id, ownerId, taskId, done, syncState",
      reminders: "id, ownerId, remindAt, targetType, targetId, enabled, syncState",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      flashcardReviews: "id, ownerId, flashcardId, subjectId, reviewedAt, syncState",
      quizzes: "id, subjectId, completedAt",
      quizAttempts: "id, ownerId, quizId, subjectId, submittedAt, score, syncState",
      documents: "id, type, status, dueAt",
      documentVersions: "id, ownerId, documentId, createdAt, syncState",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      agentMemories: "id, ownerId, category, groupId, enabled, updatedAt",
      agentConversations: "id, ownerId, scope, groupId, updatedAt",
      agentMessages: "id, conversationId, role, createdAt",
      auditLog: "id, ownerId, action, impact, createdAt, affectedType, affectedId",
      actionHistory: "id, ownerId, action, status, createdAt, expiresAt",
      settings: "id",
    });
    this.version(6).stores({
      accounts: "id, email, profileId, mode, lastSignedInAt",
      authSessions: "id, accountId, profileId, active, updatedAt",
      onboarding: "id, accountId, profileId, completed, updatedAt",
      profile: "id",
      subjects: "id, name, code",
      sessions: "id, subjectId, weekday",
      calendarEvents: "id, ownerId, startAt, subjectId, taskId, groupId, syncState",
      tasks: "id, subjectId, status, dueAt, groupId",
      checklistItems: "id, ownerId, taskId, done, syncState",
      reminders: "id, ownerId, remindAt, targetType, targetId, enabled, syncState",
      materials: "id, subjectId, type, folder, updatedAt",
      highlights: "id, materialId, createdAt",
      flashcards: "id, deck, subjectId, dueAt",
      flashcardReviews: "id, ownerId, flashcardId, subjectId, reviewedAt, syncState",
      quizzes: "id, subjectId, completedAt",
      quizAttempts: "id, ownerId, quizId, subjectId, submittedAt, score, syncState",
      documents: "id, type, status, dueAt",
      documentVersions: "id, ownerId, documentId, createdAt, syncState",
      groups: "id, subjectId",
      chats: "id, groupId, createdAt",
      studySessions: "id, subjectId, startedAt",
      agentMemories: "id, ownerId, category, groupId, enabled, updatedAt",
      agentConversations: "id, ownerId, scope, groupId, updatedAt",
      agentMessages: "id, conversationId, role, createdAt",
      auditLog: "id, ownerId, action, impact, createdAt, affectedType, affectedId",
      actionHistory: "id, ownerId, action, status, createdAt, expiresAt",
      settings: "id",
    });
  }
}

export const lumaDb = new LumaDatabase();

export const SEED_VERSION_STORAGE_KEY = "luma.seedVersion";
export const CURRENT_OWNER_STORAGE_KEY = "luma.currentOwnerId";
export const DEFAULT_OWNER_ID = "profile-ava";

export function getCurrentOwnerId() {
  try {
    return globalThis.localStorage?.getItem(CURRENT_OWNER_STORAGE_KEY) ?? DEFAULT_OWNER_ID;
  } catch {
    return DEFAULT_OWNER_ID;
  }
}

export function setCurrentOwnerId(ownerId: string) {
  try {
    globalThis.localStorage?.setItem(CURRENT_OWNER_STORAGE_KEY, ownerId);
  } catch {
    // IndexedDB remains the source of truth when localStorage is unavailable.
  }
}

export async function getActiveAuthSession() {
  await seedDatabase();
  return (await lumaDb.authSessions.toArray()).find((session) => session.active);
}

function getStoredSeedVersion() {
  try {
    return globalThis.localStorage?.getItem(SEED_VERSION_STORAGE_KEY) ?? null;
  } catch {
    return null;
  }
}

function setStoredSeedVersion(version: string) {
  try {
    globalThis.localStorage?.setItem(SEED_VERSION_STORAGE_KEY, version);
  } catch {
    // Storage can be unavailable in restricted browser contexts; IndexedDB still receives the seed.
  }
}

async function clearAllTables() {
  await Promise.all([
    lumaDb.accounts.clear(),
    lumaDb.authSessions.clear(),
    lumaDb.onboarding.clear(),
    lumaDb.profile.clear(),
    lumaDb.subjects.clear(),
    lumaDb.sessions.clear(),
    lumaDb.calendarEvents.clear(),
    lumaDb.tasks.clear(),
    lumaDb.checklistItems.clear(),
    lumaDb.reminders.clear(),
    lumaDb.materials.clear(),
    lumaDb.highlights.clear(),
    lumaDb.flashcards.clear(),
    lumaDb.flashcardReviews.clear(),
    lumaDb.quizzes.clear(),
    lumaDb.quizAttempts.clear(),
    lumaDb.documents.clear(),
    lumaDb.documentVersions.clear(),
    lumaDb.groups.clear(),
    lumaDb.chats.clear(),
    lumaDb.studySessions.clear(),
    lumaDb.agentMemories.clear(),
    lumaDb.agentConversations.clear(),
    lumaDb.agentMessages.clear(),
    lumaDb.auditLog.clear(),
    lumaDb.actionHistory.clear(),
    lumaDb.settings.clear(),
  ]);
}

export async function seedDatabase(force = false) {
  const existing = await lumaDb.profile.count();
  const existingAccounts = await lumaDb.accounts.count();
  const seedIsCurrent = getStoredSeedVersion() === TIMETABLE_SEED_VERSION;
  if (existing > 0 && existingAccounts > 0 && seedIsCurrent && !force) {
    return;
  }

  const seed = createSeedData();
  await lumaDb.transaction(
    "rw",
    [
      lumaDb.profile,
      lumaDb.subjects,
      lumaDb.sessions,
      lumaDb.calendarEvents,
      lumaDb.tasks,
      lumaDb.checklistItems,
      lumaDb.reminders,
      lumaDb.materials,
      lumaDb.highlights,
      lumaDb.flashcards,
      lumaDb.flashcardReviews,
      lumaDb.quizzes,
      lumaDb.quizAttempts,
      lumaDb.documents,
      lumaDb.documentVersions,
      lumaDb.groups,
      lumaDb.chats,
      lumaDb.studySessions,
      lumaDb.agentMemories,
      lumaDb.agentConversations,
      lumaDb.agentMessages,
      lumaDb.auditLog,
      lumaDb.actionHistory,
      lumaDb.settings,
      lumaDb.accounts,
      lumaDb.authSessions,
      lumaDb.onboarding,
    ],
    async () => {
      if (force || !seedIsCurrent) {
        await clearAllTables();
      }
      await lumaDb.accounts.put(seed.account);
      await lumaDb.authSessions.put(seed.authSession);
      await lumaDb.onboarding.put(seed.onboarding);
      await lumaDb.profile.put(seed.profile);
      await lumaDb.subjects.bulkPut(seed.subjects);
      await lumaDb.sessions.bulkPut(seed.sessions);
      await lumaDb.calendarEvents.bulkPut(seed.calendarEvents);
      await lumaDb.tasks.bulkPut(seed.tasks);
      await lumaDb.checklistItems.bulkPut(seed.checklistItems);
      await lumaDb.reminders.bulkPut(seed.reminders);
      await lumaDb.materials.bulkPut(seed.materials);
      await lumaDb.highlights.bulkPut(seed.highlights);
      await lumaDb.flashcards.bulkPut(seed.flashcards);
      await lumaDb.flashcardReviews.bulkPut(seed.flashcardReviews);
      await lumaDb.quizzes.bulkPut(seed.quizzes);
      await lumaDb.quizAttempts.bulkPut(seed.quizAttempts);
      await lumaDb.documents.bulkPut(seed.documents);
      await lumaDb.documentVersions.bulkPut(seed.documentVersions);
      await lumaDb.groups.bulkPut(seed.groups);
      await lumaDb.chats.bulkPut(seed.chats);
      await lumaDb.studySessions.bulkPut(seed.studySessions);
      await lumaDb.agentMemories.bulkPut(seed.agentMemories);
      await lumaDb.agentConversations.bulkPut(seed.agentConversations);
      await lumaDb.agentMessages.bulkPut(seed.agentMessages);
      await lumaDb.auditLog.bulkPut(seed.auditLog);
      await lumaDb.actionHistory.bulkPut(seed.actionHistory);
      await lumaDb.settings.put(seed.settings);
    },
  );
  setStoredSeedVersion(TIMETABLE_SEED_VERSION);
  setCurrentOwnerId(seed.authSession.profileId);
}

export async function loadSnapshot(): Promise<LumaSnapshot> {
  await seedDatabase();
  const fallback = createSeedData();
  const activeSession = (await lumaDb.authSessions.toArray()).find((session) => session.active);
  if (!activeSession) {
    throw new Error("auth_session_required");
  }
  const authSession = activeSession;
  const account = (await lumaDb.accounts.get(authSession.accountId)) ?? fallback.account;
  const onboarding = (await lumaDb.onboarding.where("accountId").equals(account.id).first()) ?? fallback.onboarding;
  const profileId = authSession.profileId || account.profileId || fallback.profile.id;
  setCurrentOwnerId(profileId);

  const [
    profile,
    subjects,
    sessions,
    calendarEvents,
    tasks,
    checklistItems,
    reminders,
    materials,
    highlights,
    flashcards,
    flashcardReviews,
    quizzes,
    quizAttempts,
    documents,
    documentVersions,
    groups,
    chats,
    studySessions,
    agentMemories,
    agentConversations,
    agentMessages,
    auditLog,
    actionHistory,
    savedSettings,
  ] = await Promise.all([
    lumaDb.profile.get(profileId),
    lumaDb.subjects.toArray(),
    lumaDb.sessions.toArray(),
    lumaDb.calendarEvents.toArray(),
    lumaDb.tasks.toArray(),
    lumaDb.checklistItems.toArray(),
    lumaDb.reminders.toArray(),
    lumaDb.materials.toArray(),
    lumaDb.highlights.toArray(),
    lumaDb.flashcards.toArray(),
    lumaDb.flashcardReviews.toArray(),
    lumaDb.quizzes.toArray(),
    lumaDb.quizAttempts.toArray(),
    lumaDb.documents.toArray(),
    lumaDb.documentVersions.toArray(),
    lumaDb.groups.toArray(),
    lumaDb.chats.toArray(),
    lumaDb.studySessions.toArray(),
    lumaDb.agentMemories.toArray(),
    lumaDb.agentConversations.orderBy("updatedAt").toArray(),
    lumaDb.agentMessages.orderBy("createdAt").toArray(),
    lumaDb.auditLog.toArray(),
    lumaDb.actionHistory.toArray(),
    lumaDb.settings.get(profileId === DEFAULT_OWNER_ID ? "settings" : `settings-${profileId}`),
  ]);

  const legacySettings = profileId === DEFAULT_OWNER_ID ? undefined : await lumaDb.settings.get("settings");
  const settingsSource = savedSettings ?? legacySettings;
  const hydratedSettings = settingsSource ? { ...fallback.settings, ...settingsSource } : fallback.settings;
  const activeSettings = {
    ...hydratedSettings,
    id: profileId === DEFAULT_OWNER_ID ? "settings" : `settings-${profileId}`,
  };
  const normalizedProfile = profile
    ? {
        ...fallback.profile,
        ...profile,
        visibility: { ...fallback.profile.visibility, ...profile.visibility },
        currentGoal: profile.currentGoal ? { ...fallback.profile.currentGoal, ...profile.currentGoal } : fallback.profile.currentGoal,
        achievements: profile.achievements ?? fallback.profile.achievements,
        journey: profile.journey ?? fallback.profile.journey,
        preferencesSnapshot: profile.preferencesSnapshot ?? fallback.profile.preferencesSnapshot,
        metricsVisible: profile.metricsVisible ?? fallback.profile.metricsVisible,
      }
    : fallback.profile;

  const isDemoOwner = profileId === DEFAULT_OWNER_ID;
  const ownerMatches = (owner?: string) => owner === profileId || (isDemoOwner && !owner);
  const metaOwnerMatches = (meta?: { ownerId?: string }) => ownerMatches(meta?.ownerId);
  const visibleCalendarEvents = calendarEvents.filter((event) => ownerMatches(event.ownerId));
  const visibleChecklistItems = checklistItems.filter((item) => ownerMatches(item.ownerId));
  const visibleReminders = reminders.filter((reminder) => ownerMatches(reminder.ownerId));
  const visibleMaterials = materials.filter((material) => metaOwnerMatches(material.meta));
  const visibleMaterialIds = new Set(visibleMaterials.map((material) => material.id));
  const visibleHighlights = highlights.filter((highlight) => ownerMatches(highlight.ownerId) && visibleMaterialIds.has(highlight.materialId));
  const visibleFlashcards = flashcards.filter((card) => ownerMatches(card.ownerId) && (!card.sourceMaterialId || visibleMaterialIds.has(card.sourceMaterialId)));
  const visibleFlashcardIds = new Set(visibleFlashcards.map((card) => card.id));
  const visibleFlashcardReviews = flashcardReviews.filter((review) => ownerMatches(review.ownerId) && visibleFlashcardIds.has(review.flashcardId));
  const visibleQuizzes = quizzes.filter((quiz) => ownerMatches(quiz.ownerId));
  const visibleQuizIds = new Set(visibleQuizzes.map((quiz) => quiz.id));
  const visibleQuizAttempts = quizAttempts.filter((attempt) => ownerMatches(attempt.ownerId) && visibleQuizIds.has(attempt.quizId));
  const visibleDocuments = documents.filter((document) => ownerMatches(document.ownerId));
  const visibleDocumentIds = new Set(visibleDocuments.map((document) => document.id));
  const visibleDocumentVersions = documentVersions.filter((version) => ownerMatches(version.ownerId) && visibleDocumentIds.has(version.documentId));
  const visibleGroups = groups.filter((group) => ownerMatches(group.ownerId) || group.members.some((member) => member.id === profileId));
  const visibleGroupIds = new Set(visibleGroups.map((group) => group.id));
  const visibleTasks = tasks.filter((task) => metaOwnerMatches(task.meta) || (task.groupId ? visibleGroupIds.has(task.groupId) : false));
  const visibleTaskIds = new Set(visibleTasks.map((task) => task.id));
  const visibleChats = chats.filter((chat) => visibleGroupIds.has(chat.groupId) && ownerMatches(chat.ownerId));
  const visibleStudySessions = studySessions.filter((session) => metaOwnerMatches(session.meta));
  const visibleAgentMemories = agentMemories.filter((memory) => ownerMatches(memory.ownerId));
  const visibleAgentConversations = agentConversations.filter((conversation) => ownerMatches(conversation.ownerId));
  const visibleAgentConversationIds = new Set(visibleAgentConversations.map((conversation) => conversation.id));
  const visibleAgentMessages = agentMessages.filter((message) => visibleAgentConversationIds.has(message.conversationId));
  const visibleAuditLog = auditLog.filter((entry) => ownerMatches(entry.ownerId));
  const visibleActionHistory = actionHistory.filter((entry) => ownerMatches(entry.ownerId));

  return {
    account,
    authSession,
    onboarding,
    profile: normalizedProfile,
    subjects,
    sessions,
    calendarEvents: visibleCalendarEvents,
    tasks: visibleTasks,
    checklistItems: visibleChecklistItems.filter((item) => visibleTaskIds.has(item.taskId)),
    reminders: visibleReminders,
    materials: visibleMaterials,
    highlights: visibleHighlights,
    flashcards: visibleFlashcards,
    flashcardReviews: visibleFlashcardReviews,
    quizzes: visibleQuizzes,
    quizAttempts: visibleQuizAttempts,
    documents: visibleDocuments,
    documentVersions: visibleDocumentVersions,
    groups: visibleGroups,
    chats: visibleChats,
    studySessions: visibleStudySessions,
    agentMemories: visibleAgentMemories,
    agentConversations: visibleAgentConversations,
    agentMessages: visibleAgentMessages,
    auditLog: visibleAuditLog,
    actionHistory: visibleActionHistory,
    settings: activeSettings,
  };
}

export async function resetDatabase() {
  await seedDatabase(true);
  return loadSnapshot();
}
