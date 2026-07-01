export type RouteId = "home" | "learn" | "calendar" | "together" | "create" | "profile" | "settings";

export type ThemeName = "Night Bloom" | "Deep Ocean" | "Forest Mist" | "Sunset Glow" | "Monochrome";
export type ExperiencePreset = "Balanced" | "Exam Sprint" | "Project Studio" | "Minimal Focus";
export type BackgroundScene = "Night City" | "Deep Ocean" | "Forest Mist" | "Sunset Glow";
export type WidgetSize = "small" | "medium" | "large" | "hero";
export type TaskStatus = "todo" | "in-progress" | "done";
export type TaskType = "homework" | "test" | "presentation" | "project" | "revision" | "admin";
export type MaterialType = "pdf" | "image" | "doc" | "slide" | "sheet" | "link" | "note" | "text";
export type DocumentType = "Essay" | "Report" | "Presentation" | "Reflection" | "Study Guide" | "Project Plan";
export type ClassSessionMode = "regular" | "flipped" | "self-study";
export type AiProviderStatus = "ready" | "offline" | "checking" | "error";
export type ProfileVisibility = "private" | "friends" | "groups";
export type GoalType = "exam" | "assignment" | "project" | "habit" | "subject_improvement" | "custom";
export type SyncState = "local" | "queued" | "synced" | "conflict";
export type PermissionState = "owner" | "editable" | "shared-read" | "shared-edit" | "locked";
export type VisibilityState = "private" | "group" | "public-link";
export type AuditImpact = "read" | "personal" | "shared" | "destructive";
export type ProfileMetricKey =
  | "streak"
  | "focusHours"
  | "classesTracked"
  | "tasksDone"
  | "flashcardsReviewed"
  | "quizzesCompleted"
  | "groupContributions"
  | "averageDailyFocus"
  | "upcomingDeadlines";

export interface EntityMeta {
  ownerId: string;
  createdAt: string;
  updatedAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
  visibility: VisibilityState;
}

export interface StudyGoal {
  id: string;
  title: string;
  type: GoalType;
  progress: number;
  dueAt?: string;
  targetFocusMinutes?: number;
  linkedSubjectId?: string;
  linkedTaskIds?: string[];
  linkedMaterialIds?: string[];
  status: "active" | "complete" | "archived";
}

export interface Achievement {
  id: string;
  title: string;
  description: string;
  icon: string;
  unlockedAt?: string;
  progress?: number;
  maxProgress?: number;
}

export interface StudyJourneyEntry {
  id: string;
  title: string;
  date: string;
  detail: string;
  type: "goal" | "quiz" | "project" | "achievement" | "material" | "streak" | "group";
}

export interface UserProfile {
  id: string;
  name: string;
  year: string;
  avatar: string;
  goals: string[];
  focusSubjects: string[];
  streakDays: number;
  focusMinutes: number;
  taskCompletion: number;
  username?: string;
  school?: string;
  gradeOrYear?: string;
  program?: string;
  bio?: string;
  timezone?: string;
  language?: string;
  location?: string;
  pronouns?: string;
  motto?: string;
  avatarUrl?: string;
  avatarStyle?: "monogram" | "gradient" | "photo";
  avatarGradient?: string;
  avatarCrop?: {
    zoom: number;
    x: number;
    y: number;
  };
  visibility: {
    profile: ProfileVisibility;
    avatar: ProfileVisibility;
    studyFocus: ProfileVisibility;
  };
  currentGoal?: StudyGoal;
  achievements: Achievement[];
  journey: StudyJourneyEntry[];
  preferencesSnapshot: string[];
  metricsVisible: ProfileMetricKey[];
}

export interface LocalAccount {
  id: string;
  profileId: string;
  email: string;
  displayName: string;
  passwordHash: string;
  passwordSalt: string;
  mode: "demo" | "local";
  createdAt: string;
  updatedAt: string;
  lastSignedInAt?: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface AuthSession {
  id: string;
  accountId: string;
  profileId: string;
  active: boolean;
  createdAt: string;
  updatedAt: string;
  expiresAt?: string;
}

export interface OnboardingState {
  id: string;
  accountId: string;
  profileId: string;
  completed: boolean;
  skipped: boolean;
  currentStep: number;
  completedSteps: string[];
  createdAt: string;
  updatedAt: string;
}

export interface Subject {
  id: string;
  name: string;
  code: string;
  teacher: string;
  room: string;
  building: string;
  color: string;
  materials: string[];
}

export interface ClassSession {
  id: string;
  subjectId: string;
  weekday: number;
  start: string;
  end: string;
  periodLabel?: string;
  room: string;
  building: string;
  teacher: string;
  bring: string[];
  mode?: ClassSessionMode;
  note?: string;
}

export interface TaskItem {
  id: string;
  title: string;
  subjectId: string;
  status: TaskStatus;
  dueAt: string;
  priority: "low" | "medium" | "high";
  type: TaskType;
  notes: string;
  groupId?: string;
  assignee?: string;
  checklistItemIds?: string[];
  calendarEventId?: string;
  estimatedMinutes?: number;
  meta?: EntityMeta;
}

export interface ChecklistItem {
  id: string;
  ownerId: string;
  taskId: string;
  title: string;
  done: boolean;
  createdAt: string;
  updatedAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface CalendarEvent {
  id: string;
  ownerId: string;
  title: string;
  type: "class" | "task" | "focus" | "exam" | "presentation" | "personal" | "group";
  startAt: string;
  endAt: string;
  subjectId?: string;
  taskId?: string;
  groupId?: string;
  location?: string;
  notes?: string;
  createdAt: string;
  updatedAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface Reminder {
  id: string;
  ownerId: string;
  title: string;
  remindAt: string;
  targetType: "task" | "event" | "class" | "material" | "project";
  targetId: string;
  enabled: boolean;
  createdAt: string;
  updatedAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface Material {
  id: string;
  title: string;
  subjectId: string;
  type: MaterialType;
  sourceName: string;
  folder: string;
  tags: string[];
  content: string;
  createdAt: string;
  updatedAt: string;
  sharedWith: string[];
  meta?: EntityMeta;
}

export interface Highlight {
  id: string;
  ownerId?: string;
  materialId: string;
  text: string;
  note: string;
  createdAt: string;
}

export interface Flashcard {
  id: string;
  ownerId?: string;
  deck: string;
  subjectId: string;
  front: string;
  back: string;
  mastery: number;
  dueAt: string;
  sourceMaterialId?: string;
}

export interface FlashcardReview {
  id: string;
  ownerId: string;
  flashcardId: string;
  subjectId: string;
  quality: "again" | "hard" | "good" | "easy";
  previousMastery: number;
  nextMastery: number;
  reviewedAt: string;
  nextDueAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface QuizQuestion {
  id: string;
  prompt: string;
  answer: string;
  options?: string[];
  type: "multiple-choice" | "short-answer" | "matching" | "written";
  topic: string;
}

export interface Quiz {
  id: string;
  ownerId?: string;
  title: string;
  subjectId: string;
  questions: QuizQuestion[];
  score?: number;
  completedAt?: string;
  weakTopics: string[];
}

export interface QuizAttempt {
  id: string;
  ownerId: string;
  quizId: string;
  subjectId: string;
  answers: Record<string, string>;
  score: number;
  totalQuestions: number;
  correctCount: number;
  weakTopics: string[];
  startedAt: string;
  submittedAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface CreateDocument {
  id: string;
  ownerId?: string;
  type: DocumentType;
  title: string;
  outline: string[];
  body: string;
  status: "draft" | "review" | "ready" | "archived";
  dueAt?: string;
  sourceMaterialIds: string[];
  milestones: string[];
}

export interface DocumentVersion {
  id: string;
  ownerId: string;
  documentId: string;
  title: string;
  body: string;
  outline: string[];
  label: string;
  createdAt: string;
  syncState: SyncState;
  permissionState: PermissionState;
}

export interface GroupMember {
  id: string;
  name: string;
  role: string;
  avatar: string;
  online: boolean;
}

export interface GroupWorkspace {
  id: string;
  ownerId?: string;
  name: string;
  subjectId: string;
  members: GroupMember[];
  taskIds: string[];
  materialIds: string[];
  progress: number;
  milestones: { id: string; label: string; dueAt: string; done: boolean }[];
}

export interface ChatMessage {
  id: string;
  ownerId?: string;
  groupId: string;
  author: string;
  message: string;
  createdAt: string;
}

export interface StudySession {
  id: string;
  subjectId: string;
  minutes: number;
  mode: "deep-work" | "revision" | "quiz" | "group";
  startedAt: string;
  completed: boolean;
  taskId?: string;
  endedAt?: string;
  meta?: EntityMeta;
}

export interface WidgetConfig {
  id: string;
  type: "brief" | "focus" | "vault" | "due" | "analytics" | "create" | "together";
  title: string;
  size: WidgetSize;
  priority: number;
  pinned?: string;
}

export interface UserSettings {
  id: string;
  theme: ThemeName;
  blur: number;
  widgetStyle: "Glassmorphism" | "Soft Glow" | "Minimal";
  accentColor: string;
  language: "English" | "Thai" | "Japanese" | "System Default";
  timeFormat: "12-hour" | "24-hour";
  weekStart: "Monday" | "Sunday" | "Saturday";
  landingPage: "Home" | "Calendar" | "Learn" | "Last Opened Page";
  lastOpenedRoute: RouteId;
  autoStart: boolean;
  minimizeToDock: boolean;
  soundEffects: boolean;
  focusOnLaunch: boolean;
  density: "Compact" | "Comfortable" | "Spacious";
  motion: "low" | "balanced" | "expressive";
  focusSound: "Rain" | "Ocean" | "Forest" | "None";
  focusSessionMinutes: number;
  shortBreakMinutes: number;
  longBreakMinutes: number;
  focusAutoStartBreaks: boolean;
  focusAutoStartSessions: boolean;
  notificationsEnabled: boolean;
  deadlineReminders: boolean;
  dailyAgenda: boolean;
  quietHoursEnabled: boolean;
  quietHoursStart: string;
  quietHoursEnd: string;
  localAiEndpoint: string;
  localAiModel: string;
  localAiPerformance: "Balanced" | "Battery Saver" | "Fastest" | "Best Quality";
  localOnlyMaterials: boolean;
  allowCloudSync: boolean;
  aiMemoryEnabled: boolean;
  diagnostics: boolean;
  groupAiAccess: boolean;
  widgetsSnapToGrid: boolean;
  showWidgetLabels: boolean;
  backupFrequency: "Daily" | "Weekly" | "Manual";
  experiencePreset: ExperiencePreset;
  backgroundScene: BackgroundScene;
  homeLayout: WidgetConfig[];
}

export interface AgentAction {
  id: string;
  type:
    | "create-task"
    | "reschedule-task"
    | "create-flashcards"
    | "create-quiz"
    | "create-outline"
    | "share-material"
    | "start-focus"
    | "open-route";
  label: string;
  style?: "primary" | "secondary" | "danger";
  preview?: {
    title: string;
    summary: string;
    steps: string[];
    impact: AuditImpact;
    sourceRefs?: SourceReference[];
  };
  requiresConfirmation?: boolean;
  confirmation?: {
    title: string;
    description: string;
    impact: AuditImpact;
  };
  payload: Record<string, string | number | boolean | string[]>;
}

export interface SourceReference {
  type: "file" | "task" | "calendar" | "group" | "document" | "memory" | "settings" | "focus";
  id: string;
  title: string;
  location?: string;
}

export interface UndoAction {
  label: string;
  actionId: string;
  expiresAt?: string;
}

export interface AgentToolResult<T = unknown> {
  success: boolean;
  data?: T;
  summary: string;
  sourceRefs?: SourceReference[];
  undo?: UndoAction;
  requiresConfirmation?: boolean;
  confirmation?: {
    title: string;
    description: string;
    impact: AuditImpact;
  };
  error?: {
    code: string;
    message: string;
    recoveryActions?: string[];
  };
}

export interface AgentMemory {
  id: string;
  ownerId: string;
  category: "preference" | "learning" | "workflow" | "group";
  content: string;
  enabled: boolean;
  groupId?: string;
  createdAt: string;
  updatedAt: string;
}

export interface AuditLogEntry {
  id: string;
  ownerId: string;
  action: string;
  summary: string;
  impact: AuditImpact;
  sourceRefs: SourceReference[];
  affectedType?: SourceReference["type"];
  affectedId?: string;
  provider: "local" | "ollama" | "cloud" | "manual";
  createdAt: string;
  undoActionId?: string;
}

export interface ActionHistoryEntry {
  id: string;
  ownerId: string;
  action: string;
  status: "completed" | "undone" | "failed";
  summary: string;
  payload: Record<string, unknown>;
  inversePayload?: Record<string, unknown>;
  createdAt: string;
  updatedAt: string;
  expiresAt?: string;
}

export interface AgentMessage {
  id: string;
  conversationId: string;
  role: "user" | "assistant" | "tool";
  content: string;
  createdAt: string;
  contextRefs?: SourceReference[];
  actionIds?: string[];
}

export interface AgentConversation {
  id: string;
  ownerId: string;
  scope: "personal" | "group";
  groupId?: string;
  createdAt: string;
  updatedAt: string;
}

export interface LumaSnapshot {
  account: LocalAccount;
  authSession: AuthSession;
  onboarding: OnboardingState;
  profile: UserProfile;
  subjects: Subject[];
  sessions: ClassSession[];
  calendarEvents: CalendarEvent[];
  tasks: TaskItem[];
  checklistItems: ChecklistItem[];
  reminders: Reminder[];
  materials: Material[];
  highlights: Highlight[];
  flashcards: Flashcard[];
  flashcardReviews: FlashcardReview[];
  quizzes: Quiz[];
  quizAttempts: QuizAttempt[];
  documents: CreateDocument[];
  documentVersions: DocumentVersion[];
  groups: GroupWorkspace[];
  chats: ChatMessage[];
  studySessions: StudySession[];
  agentMemories: AgentMemory[];
  agentConversations: AgentConversation[];
  agentMessages: AgentMessage[];
  auditLog: AuditLogEntry[];
  actionHistory: ActionHistoryEntry[];
  settings: UserSettings;
}

export interface AgentResult {
  answer: string;
  actions: AgentAction[];
  provider: "ollama" | "fallback" | "local";
  mode?: "ask" | "action" | "study" | "create" | "group";
  contextRefs?: SourceReference[];
  contextChips?: string[];
  undo?: UndoAction;
  status?: "idle" | "thinking" | "preview" | "completed" | "error" | "offline";
}
