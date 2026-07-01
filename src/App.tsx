import {
  DndContext,
  KeyboardSensor,
  PointerSensor,
  closestCenter,
  useSensor,
  useSensors,
  type DragEndEvent,
} from "@dnd-kit/core";
import {
  SortableContext,
  arrayMove,
  rectSortingStrategy,
  sortableKeyboardCoordinates,
  useSortable,
} from "@dnd-kit/sortable";
import { CSS } from "@dnd-kit/utilities";
import {
  addDays,
  differenceInCalendarDays,
  eachDayOfInterval,
  endOfMonth,
  endOfWeek,
  format,
  isSameDay,
  parseISO,
  startOfMonth,
  startOfWeek,
} from "date-fns";
import { Document, Packer, Paragraph, TextRun } from "docx";
import { motion } from "framer-motion";
import {
  BarChart3,
  Bell,
  BookOpen,
  Bot,
  CalendarDays,
  Check,
  ChevronLeft,
  ChevronRight,
  Circle,
  Cloud,
  Clock,
  Download,
  FilePlus2,
  FileText,
  Flame,
  Gauge,
  Globe2,
  HardDrive,
  Headphones,
  Home,
  Info,
  Languages,
  LayoutGrid,
  Lock,
  MapPin,
  Menu,
  Mic,
  Minus,
  PanelTop,
  Palette,
  Play,
  Plus,
  RefreshCw,
  RotateCcw,
  Search,
  Send,
  Settings,
  Shield,
  SlidersHorizontal,
  Smartphone,
  Share2,
  Sparkles,
  Timer,
  Trash2,
  Trophy,
  Upload,
  Volume2,
  Users,
  Wand2,
  X,
  Zap,
  type LucideIcon,
} from "lucide-react";
import pptxgen from "pptxgenjs";
import { Fragment, useEffect, useRef, useState, type CSSProperties } from "react";
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { FallbackProvider, OllamaProvider, runLumaAgent } from "./lib/ai";
import {
  BROWSER_LOCAL_MODELS,
  checkBrowserLocalSupport,
  checkBrowserModelCached,
  deleteBrowserModelCache,
  getBrowserLocalRuntimeState,
  initializeBrowserLocalModel,
  unloadBrowserLocalModel,
  type BrowserLocalSupport,
} from "./lib/browserLocalAi";
import {
  createLocalAccount,
  completeOnboarding,
  resetLocalPassword,
  signInLocalAccount,
  signOutLocalAccount,
  skipOnboarding,
} from "./lib/auth";
import {
  createDocumentTool,
  createFlashcardsTool,
  createHighlightTool,
  createGroupTool,
  createChecklistItemTool,
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
  updateDocumentTool,
  updateChecklistItemTool,
  updateGroupMemberRoleTool,
  updateGroupMessageTool,
  updateMaterialTool,
  updateTaskTool,
  updateTaskStatusTool,
  unshareMaterialFromGroupTool,
} from "./lib/actions";
import { createSeedData } from "./data/seed";
import { DEFAULT_OWNER_ID, getActiveAuthSession, getCurrentOwnerId, lumaDb, loadSnapshot, resetDatabase } from "./lib/db";
import { materialFromFile } from "./lib/files";
import { formatDue, getNextClass, subjectById, tasksDueWithin } from "./lib/scheduling";
import type {
  AgentAction,
  AgentResult,
  ChatMessage,
  ClassSession,
  CreateDocument,
  DocumentType,
  Flashcard,
  ChecklistItem,
  LumaSnapshot,
  Material,
  ProfileMetricKey,
  Quiz,
  RouteId,
  SourceReference,
  StudyGoal,
  StudySession,
  TaskItem,
  TaskType,
  UserProfile,
  UserSettings,
  WidgetConfig,
  ExperiencePreset,
  BackgroundScene,
  GoalType,
} from "./lib/types";

type AgentFocusTarget = {
  route: RouteId;
  sourceType?: SourceReference["type"];
  sourceId?: string;
};

const navItems: { id: RouteId; label: string; icon: typeof Home }[] = [
  { id: "home", label: "Home", icon: Home },
  { id: "calendar", label: "Calendar", icon: CalendarDays },
  { id: "learn", label: "Learn", icon: BookOpen },
  { id: "together", label: "Together", icon: Users },
  { id: "create", label: "Create", icon: FilePlus2 },
  { id: "profile", label: "Profile", icon: Users },
];

const documentTypes: DocumentType[] = ["Essay", "Report", "Presentation", "Reflection", "Study Guide", "Project Plan"];

type CalendarView = "Day" | "Week" | "Month" | "Exam Season";

const uid = (prefix: string) => `${prefix}-${crypto.randomUUID()}`;

const weekdayLabels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

function minutesFromTime(time: string) {
  const [hours, minutes] = time.split(":").map(Number);
  return hours * 60 + minutes;
}

function decimalHour(time: string) {
  return minutesFromTime(time) / 60;
}

function sessionTone(session: ClassSession) {
  if (session.mode === "flipped") return "violet";
  if (session.mode === "self-study") return "green";
  return "blue";
}

function sessionModeLabel(session: ClassSession) {
  if (session.mode === "flipped") return "Flipped / online";
  if (session.mode === "self-study") return "Self-study";
  return "Normal class";
}

function sessionTimeLabel(session: ClassSession) {
  const period = session.periodLabel ? `${session.periodLabel} · ` : "";
  return `${period}${session.start} - ${session.end}`;
}

const backgroundScenes: { id: BackgroundScene; file: string; description: string }[] = [
  {
    id: "Night City",
    file: "/assets/backgrounds/luma-night-city.png",
    description: "The signature LUMA skyline: soft violet, calm, and balanced.",
  },
  {
    id: "Deep Ocean",
    file: "/assets/backgrounds/luma-deep-ocean.png",
    description: "Cool blue focus with high contrast for long work sessions.",
  },
  {
    id: "Forest Mist",
    file: "/assets/backgrounds/luma-forest-mist.png",
    description: "Quiet, spacious, and low-pressure for deep reading.",
  },
  {
    id: "Sunset Glow",
    file: "/assets/backgrounds/luma-sunset-glow.png",
    description: "Warm creative energy for projects and presentations.",
  },
];

function sceneImage(scene: BackgroundScene) {
  return backgroundScenes.find((item) => item.id === scene)?.file ?? backgroundScenes[0].file;
}

const starterPresets: {
  id: ExperiencePreset;
  description: string;
  theme: UserSettings["theme"];
  backgroundScene: BackgroundScene;
  widgetStyle: UserSettings["widgetStyle"];
  focusSound: UserSettings["focusSound"];
  blur: number;
  layout: Pick<WidgetConfig, "type" | "title" | "size">[];
}[] = [
  {
    id: "Balanced",
    description: "Brief, focus, and vault as the calm default LUMA playground.",
    theme: "Night Bloom",
    backgroundScene: "Night City",
    widgetStyle: "Glassmorphism",
    focusSound: "Rain",
    blur: 70,
    layout: [
      { type: "brief", title: "Today's Brief", size: "hero" },
      { type: "focus", title: "Focus", size: "medium" },
      { type: "vault", title: "Study Vault", size: "medium" },
    ],
  },
  {
    id: "Exam Sprint",
    description: "Prioritizes deadlines, flashcards, focus, and weak-topic review.",
    theme: "Deep Ocean",
    backgroundScene: "Deep Ocean",
    widgetStyle: "Soft Glow",
    focusSound: "Ocean",
    blur: 82,
    layout: [
      { type: "due", title: "Exam Countdown", size: "hero" },
      { type: "focus", title: "Focus", size: "large" },
      { type: "vault", title: "Revision Vault", size: "medium" },
      { type: "brief", title: "Today's Brief", size: "medium" },
      { type: "analytics", title: "Study Pulse", size: "small" },
    ],
  },
  {
    id: "Project Studio",
    description: "Best for reports, presentations, group work, and milestone tracking.",
    theme: "Sunset Glow",
    backgroundScene: "Sunset Glow",
    widgetStyle: "Glassmorphism",
    focusSound: "None",
    blur: 64,
    layout: [
      { type: "create", title: "Create", size: "hero" },
      { type: "together", title: "Together", size: "large" },
      { type: "due", title: "Milestones", size: "medium" },
      { type: "vault", title: "Sources", size: "medium" },
      { type: "focus", title: "Focus", size: "small" },
    ],
  },
  {
    id: "Minimal Focus",
    description: "A quiet setup with only the essentials visible.",
    theme: "Monochrome",
    backgroundScene: "Forest Mist",
    widgetStyle: "Minimal",
    focusSound: "Forest",
    blur: 48,
    layout: [
      { type: "focus", title: "Focus", size: "hero" },
      { type: "brief", title: "Today", size: "medium" },
      { type: "due", title: "Next Actions", size: "medium" },
    ],
  },
];

function applyStarterPreset(settings: UserSettings, presetId: ExperiencePreset): UserSettings {
  const preset = starterPresets.find((item) => item.id === presetId) ?? starterPresets[0];
  return {
    ...settings,
    experiencePreset: preset.id,
    theme: preset.theme,
    backgroundScene: preset.backgroundScene,
    widgetStyle: preset.widgetStyle,
    focusSound: preset.focusSound,
    blur: preset.blur,
    homeLayout: preset.layout.map((item, index) => ({
      id: `widget-${preset.id.toLowerCase().replaceAll(" ", "-")}-${item.type}-${index}`,
      type: item.type,
      title: item.title,
      size: item.size,
      priority: index + 1,
    })),
  };
}

function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

const profileAvatarGradients = [
  "linear-gradient(145deg, rgba(211, 229, 255, 0.95), rgba(147, 122, 255, 0.78) 48%, rgba(19, 28, 74, 0.98))",
  "linear-gradient(145deg, rgba(255, 234, 198, 0.96), rgba(194, 138, 255, 0.72) 46%, rgba(20, 24, 68, 0.98))",
  "linear-gradient(145deg, rgba(194, 255, 240, 0.96), rgba(108, 173, 255, 0.72) 48%, rgba(16, 27, 76, 0.98))",
  "linear-gradient(145deg, rgba(255, 206, 224, 0.96), rgba(141, 112, 255, 0.76) 46%, rgba(21, 26, 71, 0.99))",
];

const profileGoalTypeOptions: { value: GoalType; label: string }[] = [
  { value: "exam", label: "Prepare for Exam" },
  { value: "assignment", label: "Finish Assignment" },
  { value: "project", label: "Complete Project" },
  { value: "habit", label: "Build a Habit" },
  { value: "subject_improvement", label: "Improve a Subject" },
  { value: "custom", label: "Custom Goal" },
];

const profileMetricDefinitions: { key: ProfileMetricKey; label: string }[] = [
  { key: "streak", label: "Current Streak" },
  { key: "focusHours", label: "Focus Time" },
  { key: "classesTracked", label: "Classes Tracked" },
  { key: "tasksDone", label: "Tasks Done" },
  { key: "flashcardsReviewed", label: "Flashcards Reviewed" },
  { key: "quizzesCompleted", label: "Quizzes Completed" },
  { key: "groupContributions", label: "Group Tasks" },
  { key: "averageDailyFocus", label: "Average Daily Focus" },
  { key: "upcomingDeadlines", label: "Upcoming Deadlines" },
];

const chartTooltipStyle = {
  background: "#151b42",
  border: "1px solid rgba(255,255,255,.15)",
  color: "#fff",
  borderRadius: "14px",
};

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

function initialsForName(name: string) {
  return name
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase() ?? "")
    .join("") || "LU";
}

function formatProfileTime(timezone?: string) {
  try {
    return new Intl.DateTimeFormat("en-US", {
      hour: "numeric",
      minute: "2-digit",
      hour12: true,
      timeZone: timezone || Intl.DateTimeFormat().resolvedOptions().timeZone,
    }).format(new Date());
  } catch {
    return format(new Date(), "p");
  }
}

function setControlledInputValue(input: HTMLInputElement, value: string) {
  const descriptor = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value");
  descriptor?.set?.call(input, value);
  input.dispatchEvent(new Event("input", { bubbles: true }));
}

function queueGlobalAgentPrompt(prompt: string, autoSubmit = false) {
  const input = document.querySelector<HTMLInputElement>(".global-agent input");
  if (!input) return;
  setControlledInputValue(input, prompt);
  input.focus();
  if (autoSubmit) {
    window.setTimeout(() => input.form?.requestSubmit(), 80);
  }
}

function buildProfileDraft(profile: UserProfile): UserProfile {
  return {
    ...profile,
    currentGoal: profile.currentGoal ? { ...profile.currentGoal } : undefined,
    visibility: { ...profile.visibility },
    achievements: [...profile.achievements],
    journey: [...profile.journey],
    preferencesSnapshot: [...profile.preferencesSnapshot],
    metricsVisible: [...profile.metricsVisible],
    goals: [...profile.goals],
    focusSubjects: [...profile.focusSubjects],
  };
}

function moveItem(list: string[], from: number, to: number) {
  const next = [...list];
  const [item] = next.splice(from, 1);
  next.splice(to, 0, item);
  return next;
}

function useLumaApp() {
  const [snapshot, setSnapshot] = useState<LumaSnapshot | null>(null);
  const [loading, setLoading] = useState(true);
  const [authRequired, setAuthRequired] = useState(false);

  const refresh = async () => {
    const activeSession = await getActiveAuthSession();
    if (!activeSession) {
      setSnapshot(null);
      setAuthRequired(true);
      setLoading(false);
      throw new Error("auth_session_required");
    }
    const next = await loadSnapshot();
    setSnapshot(next);
    setAuthRequired(false);
    setLoading(false);
    return next;
  };

  useEffect(() => {
    let cancelled = false;
    const startedAt = Date.now();
    void getActiveAuthSession().then(async (session) => {
      if (!session) {
        await waitForEntryWelcome(startedAt);
        if (!cancelled) {
          setSnapshot(null);
          setAuthRequired(true);
          setLoading(false);
        }
        return;
      }
      const next = await loadSnapshot();
      await waitForEntryWelcome(startedAt);
      if (!cancelled) {
        setSnapshot(next);
        setAuthRequired(false);
        setLoading(false);
      }
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const saveSettings = async (settings: UserSettings) => {
    const ownerId = getCurrentOwnerId();
    const scopedSettings = {
      ...settings,
      id: ownerId === DEFAULT_OWNER_ID ? "settings" : `settings-${ownerId}`,
    };
    await lumaDb.settings.put(scopedSettings);
    setSnapshot((current) => (current ? { ...current, settings: scopedSettings } : current));
  };

  const saveProfile = async (profile: UserProfile) => {
    await lumaDb.profile.put(profile);
    setSnapshot((current) => (current ? { ...current, profile } : current));
  };

  const saveTasks = async (tasks: TaskItem[]) => {
    await lumaDb.tasks.bulkPut(tasks);
    setSnapshot((current) => (current ? { ...current, tasks } : current));
  };

  const saveMaterials = async (materials: Material[]) => {
    await lumaDb.materials.bulkPut(materials);
    setSnapshot((current) => (current ? { ...current, materials } : current));
  };

  const saveFlashcards = async (flashcards: Flashcard[]) => {
    await lumaDb.flashcards.bulkPut(flashcards);
    setSnapshot((current) => (current ? { ...current, flashcards } : current));
  };

  const saveQuizzes = async (quizzes: Quiz[]) => {
    await lumaDb.quizzes.bulkPut(quizzes);
    setSnapshot((current) => (current ? { ...current, quizzes } : current));
  };

  const saveDocuments = async (documents: CreateDocument[]) => {
    await lumaDb.documents.bulkPut(documents);
    setSnapshot((current) => (current ? { ...current, documents } : current));
  };

  const saveStudySessions = async (studySessions: StudySession[]) => {
    await lumaDb.studySessions.bulkPut(studySessions);
    setSnapshot((current) => (current ? { ...current, studySessions } : current));
  };

  return {
    snapshot,
    loading,
    authRequired,
    refresh,
    saveProfile,
    saveSettings,
    saveTasks,
    saveMaterials,
    saveFlashcards,
    saveQuizzes,
    saveDocuments,
    saveStudySessions,
    setSnapshot,
    setAuthRequired,
  };
}

function AuthScreen({ onAuthenticated }: { onAuthenticated: () => Promise<void> }) {
  const [mode, setMode] = useState<"signin" | "signup" | "reset">("signin");
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [form, setForm] = useState({
    name: "",
    email: "",
    password: "",
    school: "",
    gradeOrYear: "",
    program: "",
  });

  const patch = (key: keyof typeof form, value: string) => setForm((current) => ({ ...current, [key]: value }));

  const submit = async (event: React.FormEvent) => {
    event.preventDefault();
    setBusy(true);
    setNotice("");
    try {
      if (mode === "signin") {
        const result = await signInLocalAccount(form.email, form.password);
        if (!result.success) {
          setNotice(result.message);
          return;
        }
        await onAuthenticated();
        return;
      }
      if (mode === "signup") {
        const result = await createLocalAccount(form);
        if (!result.success) {
          setNotice(result.message);
          return;
        }
        await onAuthenticated();
        return;
      }
      const result = await resetLocalPassword(form.email, form.password);
      setNotice(result.message);
      if (result.success) setMode("signin");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="auth-shell theme-night-bloom">
      <AmbientBackdrop />
      <main className="auth-card" aria-label="LUMA local authentication">
        <div className="auth-brand">
          <span>LUMA</span>
          <small>Quietly ahead</small>
        </div>
        <div>
          <h1>{mode === "signin" ? "Welcome back" : mode === "signup" ? "Create your local space" : "Reset local password"}</h1>
          <p>{mode === "signup" ? "Your workspace starts private on this device. Sync stays off until you enable it." : "Use a local LUMA account. No cloud account is required."}</p>
        </div>
        <form className="auth-form" onSubmit={submit}>
          {mode === "signup" && (
            <label>
              <span>Name</span>
              <input value={form.name} onChange={(event) => patch("name", event.currentTarget.value)} placeholder="Your name" required />
            </label>
          )}
          <label>
            <span>Email</span>
            <input value={form.email} onChange={(event) => patch("email", event.currentTarget.value)} placeholder="you@school.edu" type="email" required />
          </label>
          <label>
            <span>{mode === "reset" ? "New password" : "Password"}</span>
            <input value={form.password} onChange={(event) => patch("password", event.currentTarget.value)} placeholder="At least 8 characters" type="password" required />
          </label>
          {mode === "signup" && (
            <div className="auth-grid">
              <label><span>School</span><input value={form.school} onChange={(event) => patch("school", event.currentTarget.value)} placeholder="School" /></label>
              <label><span>Grade</span><input value={form.gradeOrYear} onChange={(event) => patch("gradeOrYear", event.currentTarget.value)} placeholder="M.5" /></label>
              <label><span>Program</span><input value={form.program} onChange={(event) => patch("program", event.currentTarget.value)} placeholder="Program" /></label>
            </div>
          )}
          {notice && <p className="auth-notice" role="status">{notice}</p>}
          <button className="modal-submit-btn" disabled={busy} type="submit">
            {busy ? "Working..." : mode === "signin" ? "Sign In" : mode === "signup" ? "Create Account" : "Reset Password"}
          </button>
        </form>
        <div className="auth-switcher" aria-label="Authentication options">
          <button className={mode === "signin" ? "active" : ""} onClick={() => setMode("signin")} type="button">Sign in</button>
          <button className={mode === "signup" ? "active" : ""} onClick={() => setMode("signup")} type="button">Create account</button>
          <button className={mode === "reset" ? "active" : ""} onClick={() => setMode("reset")} type="button">Reset password</button>
        </div>
      </main>
    </div>
  );
}

function OnboardingScreen({
  snapshot,
  saveProfile,
  saveSettings,
  refresh,
}: {
  snapshot: LumaSnapshot;
  saveProfile: (profile: UserProfile) => Promise<void>;
  saveSettings: (settings: UserSettings) => Promise<void>;
  refresh: () => Promise<LumaSnapshot>;
}) {
  const [draft, setDraft] = useState({
    name: snapshot.profile.name,
    school: snapshot.profile.school ?? "",
    gradeOrYear: snapshot.profile.gradeOrYear ?? "",
    program: snapshot.profile.program ?? "",
    theme: snapshot.settings.theme,
    localOnly: snapshot.settings.localOnlyMaterials,
    memory: snapshot.settings.aiMemoryEnabled,
  });
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const firstSubject = snapshot.subjects[0]?.name ?? "your first subject";

  const finish = async (skip = false) => {
    setBusy(true);
    setNotice("");
    try {
      if (skip) {
        await skipOnboarding(snapshot.account.id);
        await refresh();
        return;
      }
      const updatedProfile: UserProfile = {
        ...snapshot.profile,
        name: draft.name.trim() || snapshot.profile.name,
        avatar: initialsForName(draft.name || snapshot.profile.name),
        school: draft.school.trim(),
        gradeOrYear: draft.gradeOrYear.trim(),
        program: draft.program.trim(),
        year: [draft.gradeOrYear.trim(), draft.program.trim()].filter(Boolean).join(" · ") || snapshot.profile.year,
      };
      await saveProfile(updatedProfile);
      await saveSettings({
        ...snapshot.settings,
        theme: draft.theme,
        localOnlyMaterials: draft.localOnly,
        aiMemoryEnabled: draft.memory,
      });
      await completeOnboarding(snapshot.account.id, ["profile", "school", "subjects", "timetable", "theme", "privacy"]);
      await refresh();
    } catch {
      setNotice("LUMA could not save onboarding. Your existing workspace is unchanged.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="auth-shell onboarding-shell theme-night-bloom">
      <AmbientBackdrop />
      <main className="auth-card onboarding-card" aria-label="LUMA first-use onboarding">
        <div className="auth-brand"><span>LUMA</span><small>First setup</small></div>
        <h1>Make this space yours</h1>
        <p>Setup stays local. You can skip now and complete details later from Profile and Settings.</p>
        <div className="onboarding-steps" aria-label="Onboarding progress">
          {["Profile", "School", "Subjects", "Timetable", "Privacy"].map((step, index) => (
            <span key={step} className={index < 2 ? "active" : ""}>{step}</span>
          ))}
        </div>
        <form className="auth-form" onSubmit={(event) => { event.preventDefault(); void finish(false); }}>
          <div className="auth-grid">
            <label><span>Name</span><input value={draft.name} onChange={(event) => setDraft((current) => ({ ...current, name: event.currentTarget.value }))} /></label>
            <label><span>School</span><input value={draft.school} onChange={(event) => setDraft((current) => ({ ...current, school: event.currentTarget.value }))} /></label>
            <label><span>Grade / year</span><input value={draft.gradeOrYear} onChange={(event) => setDraft((current) => ({ ...current, gradeOrYear: event.currentTarget.value }))} /></label>
            <label><span>Program</span><input value={draft.program} onChange={(event) => setDraft((current) => ({ ...current, program: event.currentTarget.value }))} /></label>
          </div>
          <div className="onboarding-summary">
            <span><BookOpen size={16} /> Timetable starts with {firstSubject}</span>
            <span><Shield size={16} /> Personal materials are private by default</span>
          </div>
          <div className="auth-grid">
            <label><span>Theme</span><select value={draft.theme} onChange={(event) => setDraft((current) => ({ ...current, theme: event.currentTarget.value as UserSettings["theme"] }))}><option>Night Bloom</option><option>Deep Ocean</option><option>Forest Mist</option><option>Sunset Glow</option><option>Monochrome</option></select></label>
            <label className="auth-check"><input checked={draft.localOnly} onChange={(event) => setDraft((current) => ({ ...current, localOnly: event.currentTarget.checked }))} type="checkbox" /> Local-only materials</label>
            <label className="auth-check"><input checked={draft.memory} onChange={(event) => setDraft((current) => ({ ...current, memory: event.currentTarget.checked }))} type="checkbox" /> AI memory enabled</label>
          </div>
          {notice && <p className="auth-notice" role="status">{notice}</p>}
          <div className="auth-actions">
            <button className="modal-submit-btn" disabled={busy} type="submit">{busy ? "Saving..." : "Finish Setup"}</button>
            <button className="quiet-link" disabled={busy} onClick={() => void finish(true)} type="button">Skip for now</button>
          </div>
        </form>
      </main>
    </div>
  );
}

export default function App() {
  const app = useLumaApp();
  const [route, setRoute] = useState<RouteId | null>(null);
  const [agentResult, setAgentResult] = useState<AgentResult | null>(null);
  const [navOpen, setNavOpen] = useState(false);
  const [focusTarget, setFocusTarget] = useState<AgentFocusTarget | null>(null);

  if (app.loading || !app.snapshot) {
    if (!app.loading && app.authRequired) {
      return <AuthScreen onAuthenticated={async () => { setRoute(null); await app.refresh(); }} />;
    }
    return <LoadingExperience />;
  }
  const snapshot = app.snapshot;
  if (!snapshot.onboarding.completed && !snapshot.onboarding.skipped) {
    return (
      <OnboardingScreen
        snapshot={snapshot}
        saveProfile={app.saveProfile}
        saveSettings={app.saveSettings}
        refresh={app.refresh}
      />
    );
  }
  const configuredRoute =
    snapshot.settings.landingPage === "Last Opened Page"
      ? snapshot.settings.lastOpenedRoute
      : snapshot.settings.landingPage.toLowerCase() as RouteId;
  const activeRoute = route ?? configuredRoute;

  const navigate = (nextRoute: RouteId, target?: Omit<AgentFocusTarget, "route">) => {
    setRoute(nextRoute);
    setFocusTarget(target ? { route: nextRoute, ...target } : null);
    setNavOpen(false);
    if (snapshot.settings.lastOpenedRoute !== nextRoute) {
      void app.saveSettings({ ...snapshot.settings, lastOpenedRoute: nextRoute });
    }
  };

  const executeAgentAction = async (action: AgentAction) => {
    const current = app.snapshot;
    if (!current) return;
    const latest = current;
    if (action.type === "create-task") {
      const result = await createTaskTool({
        title: String(action.payload.title ?? "New study task"),
        subjectId: String(action.payload.subjectId ?? latest.subjects[0].id),
        dueAt: String(action.payload.dueAt ?? new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString()),
        priority: String(action.payload.priority ?? "medium") as TaskItem["priority"],
        type: String(action.payload.type ?? "homework") as TaskType,
        notes: "Created by LUMA Agent after preview.",
      });
      await app.refresh();
      navigate("calendar", result.data ? { sourceType: "task", sourceId: result.data.id } : undefined);
      setAgentResult({
        answer: result.success ? `${result.summary} I opened Calendar so you can review it.` : result.error?.message ?? "I could not create that task.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: agentResult?.provider ?? "fallback",
      });
    }
    if (action.type === "reschedule-task") {
      const taskId = String(action.payload.taskId);
      const result = await moveTaskTool({
        taskId,
        naturalLanguageDate: String(action.payload.input ?? "tomorrow"),
        provider: agentResult?.provider === "ollama" ? "ollama" : "local",
      });
      await app.refresh();
      navigate("calendar", { sourceType: "task", sourceId: taskId });
      setAgentResult({
        answer: result.success ? `${result.summary} Your brief is updated.` : result.error?.message ?? "I could not move that task.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: agentResult?.provider ?? "fallback",
      });
    }
    if (action.type === "create-flashcards") {
      const material = latest.materials.find((item) => item.id === action.payload.materialId) ?? latest.materials[0];
      if (!material) {
        navigate("learn");
        setAgentResult({ answer: "Upload or create a study note first, then I can make flashcards from it.", actions: [], provider: "fallback" });
        return;
      }
      const generated = await new FallbackProvider().generateFlashcards(material.content, material.subjectId);
      const result = await createFlashcardsTool({ materialId: material.id, flashcards: generated });
      await app.refresh();
      navigate("learn", { sourceType: "file", sourceId: material.id });
      setAgentResult({
        answer: result.success ? result.summary : result.error?.message ?? "I could not create flashcards.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: "fallback",
      });
    }
    if (action.type === "create-quiz") {
      const material = latest.materials.find((item) => item.id === action.payload.materialId) ?? latest.materials[0];
      if (!material) {
        navigate("learn");
        setAgentResult({ answer: "Upload or create a study note first, then I can generate a quiz from it.", actions: [], provider: "fallback" });
        return;
      }
      const quiz = await new FallbackProvider().generateQuiz(material.content, material.subjectId);
      const result = await createQuizTool({ materialId: material.id, quiz });
      await app.refresh();
      navigate("learn", { sourceType: "file", sourceId: material.id });
      setAgentResult({
        answer: result.success ? result.summary : result.error?.message ?? "I could not create that quiz.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: "fallback",
      });
    }
    if (action.type === "create-outline") {
      const type = String(action.payload.type ?? "Essay") as DocumentType;
      const title = String(action.payload.title ?? "Generated Study Outline");
      const outline = await new FallbackProvider().createOutline(title, type, latest.materials.map((material) => material.content));
      const result = await createDocumentTool({
        type,
        title,
        outline,
        body: outline.map((item) => `${item}\nDraft this section with evidence from your Study Vault.`).join("\n\n"),
        sourceMaterialIds: latest.materials.slice(0, 2).map((material) => material.id),
        milestones: ["Structure", "Draft", "Refine", "Export"],
      });
      await app.refresh();
      navigate("create", result.data ? { sourceType: "document", sourceId: result.data.id } : undefined);
      setAgentResult({
        answer: result.success ? `Created a ${type.toLowerCase()} outline and opened Create.` : result.error?.message ?? "I could not create that outline.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: "fallback",
      });
    }
    if (action.type === "start-focus") {
      const result = await startFocusSessionTool({
        subjectId: latest.subjects[0].id,
        minutes: Number(action.payload.minutes ?? 25),
        mode: "deep-work",
        provider: agentResult?.provider === "ollama" ? "ollama" : "local",
      });
      await app.refresh();
      navigate("profile", { sourceType: "focus", sourceId: result.data?.id });
      setAgentResult({
        answer: result.success ? `${result.summary} I will keep the next action visible.` : result.error?.message ?? "I could not start focus.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: "fallback",
      });
    }
    if (action.type === "share-material") {
      const result = await shareMaterialToGroupTool({
        materialId: String(action.payload.materialId ?? ""),
        groupId: String(action.payload.groupId ?? ""),
      });
      await app.refresh();
      navigate("together", { sourceType: "file", sourceId: String(action.payload.materialId ?? "") });
      setAgentResult({
        answer: result.success ? result.summary : result.error?.message ?? "I could not share that material.",
        actions: [],
        contextRefs: result.sourceRefs,
        undo: result.undo,
        status: result.success ? "completed" : "error",
        provider: "fallback",
      });
    }
    if (action.type === "open-route") {
      navigate(String(action.payload.route ?? "home") as RouteId, {
        sourceType: typeof action.payload.sourceType === "string" ? (action.payload.sourceType as SourceReference["type"]) : undefined,
        sourceId: typeof action.payload.sourceId === "string" ? action.payload.sourceId : undefined,
      });
      setAgentResult({
        answer: `Opened ${String(action.payload.route ?? "home")}.`,
        actions: [],
        provider: agentResult?.provider ?? "fallback",
        status: "completed",
      });
    }
  };

  const screen = {
    home: (
      <HomeScreen
        snapshot={snapshot}
        setRoute={navigate}
        saveSettings={app.saveSettings}
        refresh={app.refresh}
      />
    ),
    learn: (
      <LearnScreen
        key={`learn-${focusTarget?.route === "learn" ? focusTarget.sourceId ?? "route" : "route"}`}
        snapshot={snapshot}
        refresh={app.refresh}
        focusTarget={focusTarget?.route === "learn" ? focusTarget : null}
      />
    ),
    calendar: <CalendarScreen key={`calendar-${focusTarget?.route === "calendar" ? focusTarget.sourceId ?? "route" : "route"}`} snapshot={snapshot} refresh={app.refresh} focusTarget={focusTarget?.route === "calendar" ? focusTarget : null} />,
    together: <TogetherScreen key={`together-${focusTarget?.route === "together" ? focusTarget.sourceId ?? "route" : "route"}`} snapshot={snapshot} refresh={app.refresh} focusTarget={focusTarget?.route === "together" ? focusTarget : null} />,
    create: <CreateScreen key={`create-${focusTarget?.route === "create" ? focusTarget.sourceId ?? "route" : "route"}`} snapshot={snapshot} refresh={app.refresh} focusTarget={focusTarget?.route === "create" ? focusTarget : null} />,
    profile: <ProfileScreen snapshot={snapshot} saveProfile={app.saveProfile} navigate={navigate} />,
    settings: (
      <SettingsScreen
        snapshot={snapshot}
        saveSettings={app.saveSettings}
        reset={async () => app.setSnapshot(await resetDatabase())}
        signOut={async () => {
          await signOutLocalAccount();
          setRoute(null);
          app.setSnapshot(null);
          app.setAuthRequired(true);
        }}
      />
    ),
  }[activeRoute];

  return (
    <div
      className={`app-shell route-${activeRoute} theme-${snapshot.settings.theme.toLowerCase().replaceAll(" ", "-")}`}
      style={
        {
          "--glass-blur": `${Math.max(10, Math.round(snapshot.settings.blur / 3.8))}px`,
          "--scene-image": `url("${sceneImage(snapshot.settings.backgroundScene)}")`,
          "--accent": snapshot.settings.accentColor,
        } as CSSProperties
      }
    >
      <AmbientBackdrop />
      <SideRail route={activeRoute} setRoute={navigate} />
      <main className={`app-main ${activeRoute === "home" ? "home-main" : ""}`}>
        <motion.section
          key={activeRoute}
          className={`screen screen-${activeRoute}`}
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.3 }}
        >
          {screen}
        </motion.section>
      </main>
      <ResponsiveNav route={activeRoute} setRoute={navigate} open={navOpen} setOpen={setNavOpen} />
      <GlobalAgent
        snapshot={snapshot}
        agentResult={agentResult}
        setAgentResult={setAgentResult}
        executeAction={executeAgentAction}
        route={activeRoute}
        navigate={navigate}
        refresh={app.refresh}
      />
    </div>
  );
}

type SettingsCategory = "General" | "Appearance" | "Focus" | "Notifications" | "Widgets" | "AI Privacy" | "Local Model" | "Backup & Sync" | "About LUMA";

const settingsMenuItems: { title: SettingsCategory; description: string; icon: LucideIcon }[] = [
  { title: "General", description: "Language, time, and system", icon: SlidersHorizontal },
  { title: "Appearance", description: "Themes, colors, blur, layout", icon: Palette },
  { title: "Focus", description: "Deep work and focus mode", icon: Timer },
  { title: "Notifications", description: "Alerts, reminders, sounds", icon: Bell },
  { title: "Widgets", description: "Manage and customize widgets", icon: LayoutGrid },
  { title: "AI Privacy", description: "Privacy, memory, data control", icon: Shield },
  { title: "Local Model", description: "LUMA local model settings", icon: Bot },
  { title: "Backup & Sync", description: "Sync, backup, and restore", icon: Cloud },
  { title: "About LUMA", description: "Version, license, updates", icon: Info },
];

const accentChoices = [
  "#746bff",
  "#6f8cff",
  "#56d2e4",
  "#6ed98b",
  "#f0c24a",
  "#ee9138",
  "#db5e9d",
  "#9a75ec",
];

const widgetStyles: { value: UserSettings["widgetStyle"]; label: string; description: string }[] = [
  { value: "Glassmorphism", label: "Glass", description: "Translucent and vibrant" },
  { value: "Soft Glow", label: "Soft", description: "Subtle and smooth" },
  { value: "Minimal", label: "Solid", description: "Clean and minimal" },
];

const ENTRY_WELCOME_DURATION_MS = 1400;

function waitForEntryWelcome(startedAt: number) {
  const remaining = Math.max(0, ENTRY_WELCOME_DURATION_MS - (Date.now() - startedAt));
  return new Promise((resolve) => setTimeout(resolve, remaining));
}

function LoadingExperience() {
  return (
    <div className="loading-stage theme-night-bloom">
      <div className="loading-welcome is-minimal" role="status" aria-live="polite" aria-label="LUMA is preparing your workspace">
        <div className="loading-copy">
          <h1>Welcome back to LUMA</h1>
        </div>
        <div className="loading-progress" aria-hidden="true">
          <span />
        </div>
      </div>
    </div>
  );
}

function AmbientBackdrop() {
  return (
    <div className="ambient" aria-hidden="true">
      <div className="window-arc" />
      <div className="city-line" />
      <div className="stars" />
    </div>
  );
}

function LumaLogo() {
  return (
    <div className="luma-logo" aria-label="LUMA">
      <span>LUMA</span>
      <small>M.5 IM</small>
    </div>
  );
}

function Orb({ size = "medium" }: { size?: "small" | "medium" | "large" }) {
  return <span className={`orb orb-${size}`} aria-hidden="true" />;
}

function GlassPanel({
  children,
  className = "",
  as: Element = "section",
}: {
  children: React.ReactNode;
  className?: string;
  as?: "section" | "article" | "div";
}) {
  return <Element className={`glass-panel ${className}`}>{children}</Element>;
}

function ActionButton({
  children,
  onClick,
  tone = "primary",
  type = "button",
  title,
  disabled = false,
}: {
  children: React.ReactNode;
  onClick?: () => void;
  tone?: "primary" | "ghost" | "danger";
  type?: "button" | "submit";
  title?: string;
  disabled?: boolean;
}) {
  return (
    <button className={`action-button ${tone}`} disabled={disabled} onClick={onClick} type={type} title={title}>
      {children}
    </button>
  );
}

function SideRail({ route, setRoute }: { route: RouteId; setRoute: (route: RouteId) => void }) {
  return (
    <aside className="side-rail" aria-label="Main navigation">
      <LumaLogo />
      <nav>
        {navItems.map((item) => {
          const Icon = item.icon;
          return (
            <button
              key={item.id}
              className={route === item.id ? "active" : ""}
              data-route={item.label}
              aria-label={item.label}
              onClick={() => setRoute(item.id)}
              title={item.label}
              type="button"
            >
              <Icon size={20} />
              <span>{item.label}</span>
            </button>
          );
        })}
      </nav>
      <button className={route === "settings" ? "rail-settings active" : "rail-settings"} onClick={() => setRoute("settings")} aria-label="Settings" title="Settings" type="button">
        <Settings size={19} />
      </button>
    </aside>
  );
}

function ResponsiveNav({
  route,
  setRoute,
  open,
  setOpen,
}: {
  route: RouteId;
  setRoute: (route: RouteId) => void;
  open: boolean;
  setOpen: (open: boolean) => void;
}) {
  return (
    <>
      <button
        className="nav-hamburger"
        type="button"
        aria-label={open ? "Close navigation" : "Open navigation"}
        aria-expanded={open}
        onClick={() => setOpen(!open)}
      >
        {open ? <X size={22} /> : <Menu size={22} />}
      </button>
      {open && (
        <>
          <motion.button
            className="responsive-nav-backdrop"
            type="button"
            aria-label="Close navigation"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setOpen(false)}
          />
          <motion.aside
            className="responsive-rail"
            aria-label="Main navigation"
            initial={{ x: -124, opacity: 0 }}
            animate={{ x: 0, opacity: 1 }}
            exit={{ x: -124, opacity: 0 }}
            transition={{ type: "spring", stiffness: 360, damping: 34 }}
          >
            <LumaLogo />
            <nav>
              {navItems.map((item) => {
                const Icon = item.icon;
                return (
                  <button
                    key={item.id}
                    className={route === item.id ? "active" : ""}
                    data-route={item.label}
                    aria-label={item.label}
                    onClick={() => setRoute(item.id)}
                    title={item.label}
                    type="button"
                  >
                    <Icon size={20} />
                    <span>{item.label}</span>
                  </button>
                );
              })}
            </nav>
            <button className={route === "settings" ? "rail-settings active" : "rail-settings"} onClick={() => setRoute("settings")} aria-label="Settings" title="Settings" type="button">
              <Settings size={19} />
            </button>
          </motion.aside>
        </>
      )}
    </>
  );
}

function HomeScreen({
  snapshot,
  setRoute,
  saveSettings,
  refresh,
}: {
  snapshot: LumaSnapshot;
  setRoute: (route: RouteId) => void;
  saveSettings: (settings: UserSettings) => Promise<void>;
  refresh: () => Promise<LumaSnapshot>;
}) {
  const settings = snapshot.settings;
  const [widgetLibraryOpen, setWidgetLibraryOpen] = useState(false);
  const sensors = useSensors(
    useSensor(PointerSensor, { activationConstraint: { distance: 8 } }),
    useSensor(KeyboardSensor, { coordinateGetter: sortableKeyboardCoordinates }),
  );
  const sortedWidgets = [...settings.homeLayout].sort((a, b) => a.priority - b.priority);
  const displayName = snapshot.profile.name.split(" ")[0] || "Win";
  const next = getNextClass(snapshot.sessions);
  const subject = subjectById(snapshot.subjects, next.session.subjectId);
  const dueSoon = tasksDueWithin(snapshot.tasks, 7);
  const urgentTask = dueSoon[0];
  const recentMaterial = snapshot.materials[0];
  const visibleWidgets = sortedWidgets.slice(0, 3);
  const heroWidget = visibleWidgets.find((widget) => widget.size === "hero") ?? visibleWidgets[0];
  const sideWidgets = visibleWidgets.filter((widget) => widget.id !== heroWidget?.id).slice(0, 2);

  const onDragEnd = async (event: DragEndEvent) => {
    const { active, over } = event;
    if (!over || active.id === over.id) return;
    const oldIndex = sortedWidgets.findIndex((widget) => widget.id === active.id);
    const newIndex = sortedWidgets.findIndex((widget) => widget.id === over.id);
    const moved = arrayMove(sortedWidgets, oldIndex, newIndex).map((widget, index) => ({ ...widget, priority: index + 1 }));
    await saveSettings({ ...settings, homeLayout: moved });
  };

  const resizeWidget = async (widget: WidgetConfig) => {
    const nextSize: WidgetConfig["size"] =
      widget.size === "small" ? "medium" : widget.size === "medium" ? "large" : widget.size === "large" ? "hero" : "small";
    const nextLayout = settings.homeLayout.map((item) => (item.id === widget.id ? { ...item, size: nextSize } : item));
    await saveSettings({ ...settings, homeLayout: nextLayout });
  };

  const addWidget = async (type: WidgetConfig["type"], title: string, size: WidgetConfig["size"] = "small") => {
    const id = uid("widget");
    await saveSettings({
      ...settings,
      homeLayout: [...settings.homeLayout, { id, type, title, size, priority: settings.homeLayout.length + 1 }],
    });
    setWidgetLibraryOpen(false);
  };

  const removeWidget = async (widget: WidgetConfig) => {
    const nextLayout = settings.homeLayout
      .filter((item) => item.id !== widget.id)
      .map((item, index) => ({ ...item, priority: index + 1 }));
    await saveSettings({ ...settings, homeLayout: nextLayout });
  };

  const makeHeroWidget = async (widget: WidgetConfig) => {
    const nextLayout = settings.homeLayout
      .map(
        (item): WidgetConfig => ({
        ...item,
        size: item.id === widget.id ? "hero" : item.size === "hero" ? "medium" : item.size,
        priority: item.id === widget.id ? 0 : item.priority + 1,
        }),
      )
      .sort((a, b) => a.priority - b.priority)
      .map((item, index) => ({ ...item, priority: index + 1 }));
    await saveSettings({ ...settings, homeLayout: nextLayout });
  };

  const startFocus = async () => {
    await startFocusSessionTool({
      subjectId: snapshot.subjects[0].id,
      minutes: 25,
      mode: "deep-work",
    });
    await refresh();
  };

  return (
    <div className="home-screen">
      <section className="home-command" aria-label="LUMA home command space">
        <header className="home-command-header">
          <div>
            <h1>
              Good morning, {displayName} <Sparkles size={23} />
            </h1>
            <p>Your day. Organized. Your future. Illuminated.</p>
          </div>
          <div className="home-user-actions" aria-label="LUMA status">
            <button type="button" aria-label="Notifications">
              <Bell size={22} />
            </button>
            <span className="home-avatar" aria-label={`${displayName} profile`}>
              {snapshot.profile.avatar.slice(0, 2)}
              <i aria-hidden="true" />
            </span>
          </div>
        </header>

        <div className="home-command-canvas">
          {heroWidget && (
            <CommandHeroWidget
              widget={heroWidget}
              displayName={displayName}
              next={next}
              subjectColor={subject?.color}
              subjectName={subject?.name}
              urgentTask={urgentTask}
              recentMaterial={recentMaterial}
              setRoute={setRoute}
              startFocus={startFocus}
            />
          )}

          <div className="command-side-stack">
            {sideWidgets.map((widget) => (
              <CommandSideWidget
                key={widget.id}
                widget={widget}
                next={next}
                subjectName={subject?.name}
                urgentTask={urgentTask}
                recentMaterial={recentMaterial}
                setRoute={setRoute}
                startFocus={startFocus}
              />
            ))}
          </div>
        </div>

        <button className="home-add-widget" onClick={() => setWidgetLibraryOpen((open) => !open)} type="button">
          <Plus size={21} />
          Add Widget
        </button>

        {widgetLibraryOpen && (
          <motion.div
            className="widget-library-sheet"
            initial={{ opacity: 0, y: 18, scale: 0.98 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 12 }}
          >
            <div className="sheet-title-row">
              <div>
                <h3>Home Studio</h3>
                <p>Choose what appears on this playground. The first three panels are visible.</p>
              </div>
              <button onClick={() => setWidgetLibraryOpen(false)} type="button">
                Close
              </button>
            </div>

            <div className="home-studio-grid">
              <div className="studio-section">
                <h4>Add Widget</h4>
                <div className="library-options">
                  <button onClick={() => void addWidget("due", "Due Soon", "medium")} type="button">
                    <CalendarDays size={16} />
                    Due Soon
                  </button>
                  <button onClick={() => void addWidget("analytics", "Study Pulse", "small")} type="button">
                    <BarChart3 size={16} />
                    Study Pulse
                  </button>
                  <button onClick={() => void addWidget("create", "Create", "small")} type="button">
                    <FilePlus2 size={16} />
                    Create
                  </button>
                  <button onClick={() => void addWidget("together", "Together", "small")} type="button">
                    <Users size={16} />
                    Together
                  </button>
                </div>
              </div>

              <div className="studio-section">
                <h4>Visible Panels</h4>
                <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={onDragEnd}>
                  <SortableContext items={sortedWidgets.map((widget) => widget.id)} strategy={rectSortingStrategy}>
                    <div className="widget-token-list">
                      {sortedWidgets.map((widget) => (
                        <SortableWidgetToken
                          key={widget.id}
                          widget={widget}
                          onHero={makeHeroWidget}
                          onRemove={removeWidget}
                          onResize={resizeWidget}
                        />
                      ))}
                    </div>
                  </SortableContext>
                </DndContext>
              </div>

              <div className="studio-section">
                <h4>Preset</h4>
                <div className="compact-preset-grid">
                  {starterPresets.map((preset) => (
                    <button
                      key={preset.id}
                      className={settings.experiencePreset === preset.id ? "selected" : ""}
                      onClick={() => void saveSettings(applyStarterPreset(settings, preset.id))}
                      type="button"
                    >
                      {preset.id}
                    </button>
                  ))}
                </div>
              </div>

              <div className="studio-section">
                <h4>Ambience</h4>
                <div className="compact-scene-grid">
                  {backgroundScenes.map((scene) => (
                    <button
                      key={scene.id}
                      className={settings.backgroundScene === scene.id ? "selected" : ""}
                      onClick={() => void saveSettings({ ...settings, backgroundScene: scene.id })}
                      style={{ "--scene-thumb": `url("${scene.file}")` } as CSSProperties}
                      type="button"
                    >
                      {scene.id}
                    </button>
                  ))}
                </div>
                <label className="compact-slider">
                  Blur
                  <input
                    type="range"
                    min="20"
                    max="95"
                    value={settings.blur}
                    onChange={(event) => void saveSettings({ ...settings, blur: Number(event.currentTarget.value) })}
                  />
                </label>
              </div>
            </div>
          </motion.div>
        )}
      </section>
    </div>
  );
}

function routeForWidget(widget: WidgetConfig): RouteId {
  if (widget.type === "brief" || widget.type === "due") return "calendar";
  if (widget.type === "vault") return "learn";
  if (widget.type === "focus" || widget.type === "analytics") return "profile";
  if (widget.type === "create") return "create";
  return "together";
}

function CommandWidgetIcon({ type }: { type: WidgetConfig["type"] }) {
  if (type === "brief" || type === "due") return <CalendarDays size={22} />;
  if (type === "vault") return <BookOpen size={22} />;
  if (type === "focus") return <Timer size={22} />;
  if (type === "create") return <FilePlus2 size={22} />;
  if (type === "together") return <Users size={22} />;
  return <BarChart3 size={22} />;
}

function CommandHeroWidget({
  widget,
  displayName,
  next,
  subjectColor,
  subjectName,
  urgentTask,
  recentMaterial,
  setRoute,
  startFocus,
}: {
  widget: WidgetConfig;
  displayName: string;
  next: ReturnType<typeof getNextClass>;
  subjectColor?: string;
  subjectName?: string;
  urgentTask?: TaskItem;
  recentMaterial?: Material;
  setRoute: (route: RouteId) => void;
  startFocus: () => Promise<void>;
}) {
  if (widget.type !== "brief") {
    return (
      <motion.article
        className="command-card command-brief command-generic-hero"
        whileHover={{ y: -3 }}
        onClick={() => (widget.type === "focus" ? void startFocus() : setRoute(routeForWidget(widget)))}
        role="button"
        tabIndex={0}
        onKeyDown={(event) => {
          if (event.key === "Enter" || event.key === " ") {
            if (widget.type === "focus") void startFocus();
            else setRoute(routeForWidget(widget));
          }
        }}
      >
        <div className="command-card-head">
          <h2>
            {widget.title} <Sparkles size={18} />
          </h2>
          <span>Selected hero</span>
          <Orb size="small" />
        </div>
        <div className="generic-hero-body">
          <span>
            <CommandWidgetIcon type={widget.type} />
          </span>
          <div>
            <strong>{commandWidgetPrimary(widget, urgentTask, recentMaterial)}</strong>
            <p>{commandWidgetSecondary(widget, subjectName)}</p>
          </div>
        </div>
        <div className="brief-due-row">
          <small>LUMA keeps this panel prioritized on Home</small>
          <strong>{widget.size}</strong>
          <span>Priority {widget.priority}</span>
        </div>
      </motion.article>
    );
  }

  return (
    <motion.article
      className="command-card command-brief"
      whileHover={{ y: -3 }}
      onClick={() => setRoute("calendar")}
      role="button"
      tabIndex={0}
      onKeyDown={(event) => {
        if (event.key === "Enter" || event.key === " ") setRoute("calendar");
      }}
    >
      <div className="command-card-head">
        <h2>
          {widget.title} <Sparkles size={18} />
        </h2>
        <span>Good morning, {displayName}</span>
        <Orb size="small" />
      </div>
      <div className="brief-main-row">
        <span className="brief-subject-icon" style={{ "--subject-color": subjectColor } as CSSProperties}>
          <Sparkles size={20} />
        </span>
        <div>
          <small>Next Focus</small>
          <strong>{subjectName ?? "Deep Work"}</strong>
          <p>
            {next.session.start} - {next.session.end}
          </p>
        </div>
      </div>
      <div className="brief-meta-grid">
        <div>
          <small>Space</small>
          <span>{next.session.room}</span>
        </div>
        <div>
          <small>Guide</small>
          <span>{next.session.teacher}</span>
        </div>
      </div>
      <div className="brief-soft-row">
        <small>What to Bring</small>
        {(next.session.bring.length ? next.session.bring : ["Usual class materials"]).slice(0, 3).map((item) => (
          <span key={item}>
            <Circle size={15} />
            {item}
          </span>
        ))}
      </div>
      <div className="brief-due-row">
        <small>Due Soon</small>
        <strong>{urgentTask?.title ?? "All clear"}</strong>
        <span>{urgentTask ? formatDue(urgentTask.dueAt) : "No urgent work"}</span>
      </div>
    </motion.article>
  );
}

function CommandSideWidget({
  widget,
  subjectName,
  urgentTask,
  recentMaterial,
  setRoute,
  startFocus,
}: {
  widget: WidgetConfig;
  next: ReturnType<typeof getNextClass>;
  subjectName?: string;
  urgentTask?: TaskItem;
  recentMaterial?: Material;
  setRoute: (route: RouteId) => void;
  startFocus: () => Promise<void>;
}) {
  if (widget.type === "focus") {
    return (
      <motion.article className="command-card command-focus" whileHover={{ y: -3 }}>
        <h2>
          {widget.title} <Sparkles size={18} />
        </h2>
        <p>Deep Work</p>
        <strong>25</strong>
        <span>min</span>
        <ActionButton onClick={() => void startFocus()}>Start Focus</ActionButton>
        <div className="focus-orb-stage">
          <Orb size="large" />
        </div>
      </motion.article>
    );
  }

  if (widget.type === "vault") {
    return (
      <motion.article
        className="command-card command-vault"
        whileHover={{ y: -3 }}
        onClick={() => setRoute("learn")}
        role="button"
        tabIndex={0}
        onKeyDown={(event) => {
          if (event.key === "Enter" || event.key === " ") setRoute("learn");
        }}
      >
        <h2>
          {widget.title} <Sparkles size={18} />
        </h2>
        <p>Continue Reading</p>
        <div className="vault-current">
          <span>
            <BookOpen size={21} />
          </span>
          <div>
            <strong>{recentMaterial?.title ?? "Your study space is ready"}</strong>
            <small>{recentMaterial ? "Page 42" : "Upload your first file"}</small>
          </div>
        </div>
        <div className="vault-dots" aria-hidden="true">
          <i />
          <i />
          <i />
          <i />
          <i />
        </div>
      </motion.article>
    );
  }

  return (
    <motion.article
      className="command-card command-vault command-mini-widget"
      whileHover={{ y: -3 }}
      onClick={() => setRoute(routeForWidget(widget))}
      role="button"
      tabIndex={0}
      onKeyDown={(event) => {
        if (event.key === "Enter" || event.key === " ") setRoute(routeForWidget(widget));
      }}
    >
      <h2>
        {widget.title} <Sparkles size={18} />
      </h2>
      <p>{commandWidgetSecondary(widget, subjectName)}</p>
      <div className="vault-current">
        <span>
          <CommandWidgetIcon type={widget.type} />
        </span>
        <div>
          <strong>{commandWidgetPrimary(widget, urgentTask, recentMaterial)}</strong>
          <small>Priority {widget.priority}</small>
        </div>
      </div>
    </motion.article>
  );
}

function commandWidgetPrimary(widget: WidgetConfig, urgentTask?: TaskItem, recentMaterial?: Material) {
  if (widget.type === "due") return urgentTask?.title ?? "All clear";
  if (widget.type === "vault") return recentMaterial?.title ?? "Your study space is ready";
  if (widget.type === "focus") return "25 min";
  if (widget.type === "create") return "Start a polished draft";
  if (widget.type === "together") return "Study crew in flow";
  if (widget.type === "analytics") return "Focus pulse";
  return widget.title;
}

function commandWidgetSecondary(widget: WidgetConfig, subjectName?: string) {
  if (widget.type === "due") return "Only the next important item";
  if (widget.type === "vault") return "Continue Reading";
  if (widget.type === "focus") return "Deep Work";
  if (widget.type === "create") return "Structure ideas into work";
  if (widget.type === "together") return "Shared tasks and materials";
  if (widget.type === "analytics") return subjectName ? `${subjectName} trend` : "Study trends";
  return "Selected panel";
}

function SortableWidgetToken({
  widget,
  onHero,
  onRemove,
  onResize,
}: {
  widget: WidgetConfig;
  onHero: (widget: WidgetConfig) => Promise<void>;
  onRemove: (widget: WidgetConfig) => Promise<void>;
  onResize: (widget: WidgetConfig) => Promise<void>;
}) {
  const { attributes, listeners, setNodeRef, transform, transition, isDragging } = useSortable({ id: widget.id });
  return (
    <div
      ref={setNodeRef}
      className={`widget-token ${widget.size === "hero" ? "hero" : ""} ${isDragging ? "dragging" : ""}`}
      style={{ transform: CSS.Transform.toString(transform), transition }}
      {...attributes}
    >
      <button className="token-drag" type="button" {...listeners} aria-label={`Move ${widget.title}`}>
        <LayoutGrid size={15} />
      </button>
      <div>
        <strong>{widget.title}</strong>
        <small>
          {widget.size} · priority {widget.priority}
        </small>
      </div>
      <button onClick={() => void onHero(widget)} type="button" title="Make hero">
        <Sparkles size={15} />
      </button>
      <button onClick={() => void onResize(widget)} type="button" title="Resize">
        <PanelTop size={15} />
      </button>
      <button onClick={() => void onRemove(widget)} type="button" title="Remove">
        <Trash2 size={15} />
      </button>
    </div>
  );
}

function LearnScreen({
  snapshot,
  refresh,
  focusTarget,
}: {
  snapshot: LumaSnapshot;
  refresh: () => Promise<LumaSnapshot>;
  focusTarget?: AgentFocusTarget | null;
}) {
  const focusedMaterialId =
    focusTarget?.sourceType === "file" && focusTarget.sourceId && snapshot.materials.some((material) => material.id === focusTarget.sourceId)
      ? focusTarget.sourceId
      : "";
  const [selectedId, setSelectedId] = useState(focusedMaterialId || snapshot.materials[0]?.id || "");
  const [summary, setSummary] = useState("");
  const [query, setQuery] = useState("");
  const [studyCard, setStudyCard] = useState(snapshot.flashcards[0]?.id ?? "");
  const [notice, setNotice] = useState("");
  const [tagInput, setTagInput] = useState("");
  const [folderInput, setFolderInput] = useState("");
  const [answerShown, setAnswerShown] = useState(false);
  const [activeQuizId, setActiveQuizId] = useState(snapshot.quizzes.at(-1)?.id ?? "");
  const [quizAnswers, setQuizAnswers] = useState<Record<string, string>>({});
  const selected = snapshot.materials.find((material) => material.id === selectedId) ?? snapshot.materials[0];
  const visibleMaterials = snapshot.materials.filter((material) =>
    `${material.title} ${material.tags.join(" ")} ${material.content}`.toLowerCase().includes(query.toLowerCase()),
  );

  const uploadFile = async (file: File | undefined) => {
    if (!file) return;
    setNotice(`Processing ${file.name}...`);
    try {
      const material = await materialFromFile(file, snapshot.subjects[0].id);
      const result = await createMaterialTool(material);
      if (result.data) setSelectedId(result.data.id);
      setNotice(result.success ? result.summary : result.error?.message ?? "Upload failed.");
      await refresh();
    } catch {
      setNotice("LUMA could not read that file. Try a PDF, DOCX, TXT, spreadsheet, slide deck, or image.");
    }
  };

  const generateFlashcards = async () => {
    if (!selected) {
      setSummary("Add a note or upload a file first, then LUMA can turn it into flashcards.");
      return;
    }
    const cards = await new FallbackProvider().generateFlashcards(selected.content, selected.subjectId);
    const result = await createFlashcardsTool({ materialId: selected.id, flashcards: cards });
    if (result.data?.[0]) setStudyCard(result.data[0].id);
    setNotice(result.success ? result.summary : result.error?.message ?? "Flashcards were not created.");
    await refresh();
  };

  const generateQuiz = async () => {
    if (!selected) {
      setSummary("Add a note or upload a file first, then LUMA can generate a quiz.");
      return;
    }
    const quiz = await new FallbackProvider().generateQuiz(selected.content, selected.subjectId);
    const result = await createQuizTool({ materialId: selected.id, quiz });
    if (result.data) {
      setActiveQuizId(result.data.id);
      setQuizAnswers({});
    }
    setNotice(result.success ? result.summary : result.error?.message ?? "Quiz was not created.");
    await refresh();
  };

  const saveHighlight = async () => {
    if (!selected) return;
    const result = await createHighlightTool({
      materialId: selected.id,
      text: selected.content.slice(0, 120),
      note: "Saved from reader",
    });
    setNotice(result.success ? result.summary : result.error?.message ?? "Highlight was not saved.");
    await refresh();
  };

  const renameSelected = async () => {
    if (!selected) return;
    const title = prompt("Rename material", selected.title)?.trim();
    if (!title || title === selected.title) return;
    const result = await updateMaterialTool({ materialId: selected.id, patch: { title }, summary: `Renamed "${selected.title}" to "${title}".` });
    setNotice(result.summary);
    await refresh();
  };

  const moveSelected = async () => {
    if (!selected) return;
    const folder = folderInput.trim();
    if (!folder) {
      setNotice("Enter a folder before moving this material.");
      return;
    }
    const result = await updateMaterialTool({ materialId: selected.id, patch: { folder }, summary: `Moved "${selected.title}" to ${folder}.` });
    setFolderInput("");
    setNotice(result.summary);
    await refresh();
  };

  const addTag = async () => {
    if (!selected) return;
    const tag = tagInput.trim();
    if (!tag) return;
    const result = await updateMaterialTool({
      materialId: selected.id,
      patch: { tags: [...new Set([...selected.tags, tag])] },
      summary: `Tagged "${selected.title}" with ${tag}.`,
    });
    setTagInput("");
    setNotice(result.summary);
    await refresh();
  };

  const deleteSelected = async () => {
    if (!selected || !confirm(`Delete "${selected.title}" from Study Vault? This also removes its highlights.`)) return;
    const result = await deleteMaterialTool(selected.id);
    setSelectedId("");
    setNotice(result.summary);
    await refresh();
  };

  const shareSelected = async () => {
    if (!selected) return;
    const group = snapshot.groups[0];
    if (!group) {
      setNotice("Create a group workspace before sharing material.");
      return;
    }
    if (!confirm(`Share "${selected.title}" with ${group.name}?`)) return;
    const result = await shareMaterialToGroupTool({ materialId: selected.id, groupId: group.id });
    setNotice(result.success ? result.summary : result.error?.message ?? "Material was not shared.");
    await refresh();
  };

  const askAboutSelected = () => {
    if (!selected) {
      setNotice("Upload or select a material before asking LUMA about it.");
      return;
    }
    const input = document.querySelector<HTMLInputElement>(".global-agent input");
    input?.focus();
    if (input) {
      input.value = `Summarize "${selected.title}" and show sources.`;
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }
  };

  const activeCard = snapshot.flashcards.find((card) => card.id === studyCard) ?? snapshot.flashcards[0];
  const activeQuiz = snapshot.quizzes.find((quiz) => quiz.id === activeQuizId) ?? snapshot.quizzes.at(-1);
  const activeQuizAttempt = activeQuiz ? snapshot.quizAttempts.filter((attempt) => attempt.quizId === activeQuiz.id).at(-1) : undefined;
  const reviewCard = async (quality: "again" | "hard" | "good" | "easy") => {
    if (!activeCard) return;
    const result = await reviewFlashcardTool({ flashcardId: activeCard.id, quality });
    setAnswerShown(false);
    setNotice(result.success ? result.summary : result.error?.message ?? "Review was not saved.");
    await refresh();
  };
  const submitQuiz = async () => {
    if (!activeQuiz) {
      setNotice("Generate a quiz before submitting answers.");
      return;
    }
    const result = await submitQuizAttemptTool({ quizId: activeQuiz.id, answers: quizAnswers });
    setNotice(result.success ? result.summary : result.error?.message ?? "Quiz attempt was not saved.");
    await refresh();
  };
  const continueItems = snapshot.subjects.slice(0, 3).map((subject, index) => ({
    subject,
    progress: 0,
    next: snapshot.sessions.find((session) => session.subjectId === subject.id)?.periodLabel ?? "No material yet",
    minutes: [45, 90, 140][index] ?? 45,
  }));

  return (
    <div className="universal-screen learn-universal">
      <UniversalHeader title="Learn" subtitle="Study. Understand. Retain." avatar={snapshot.profile.avatar} />
      <div className="learn-dashboard">
        <GlassPanel className="learn-hub">
          <div className="universal-panel-title">
            <h3>Learn Hub <Sparkles size={16} /></h3>
            <div className="search-row compact-search">
              <Search size={17} />
              <input value={query} onChange={(event) => setQuery(event.currentTarget.value)} placeholder="Search topics..." />
              <button type="button" title="Search"><Search size={18} /></button>
            </div>
          </div>
          <div className="learn-section-label"><span>Continue Learning</span><small>View all</small></div>
          <div className="continue-grid">
            {continueItems.map(({ subject, progress, next, minutes }) => (
              <button key={subject.id} className="continue-card" onClick={() => setSelectedId(snapshot.materials.find((material) => material.subjectId === subject.id)?.id ?? "")} type="button">
                <span className="subject-cube" style={{ "--subject-color": subject.color } as CSSProperties}><BookOpen size={21} /></span>
                <strong>{subject.name}</strong>
                <small>{subject.code}</small>
                <i><b style={{ width: `${progress}%` }} /></i>
                <em>{progress}%</em>
                <span>{next}</span>
                <small>{minutes} min class block</small>
              </button>
            ))}
          </div>
          <div className="learn-section-label"><span>Study Tools</span></div>
          <div className="study-tools">
            {[
              ["Flashcards", "Spaced repetition", BookOpen, generateFlashcards],
              ["Quizzes", "Test your knowledge", Wand2, generateQuiz],
              ["Summaries", "AI key takeaways", FileText, () => new FallbackProvider().summarize(selected.content).then(setSummary)],
              ["AI Chat", "Ask LUMA anything", Bot, askAboutSelected],
            ].map(([label, helper, Icon, action]) => {
              const ToolIcon = Icon as typeof BookOpen;
              return (
                <button key={label as string} onClick={() => void (action as () => Promise<void> | void)()} type="button">
                  <span><ToolIcon size={18} /></span>
                  <strong>{label as string}</strong>
                  <small>{helper as string}</small>
                </button>
              );
            })}
            <label className="study-upload">
              <span><Upload size={18} /></span>
              <strong>Upload</strong>
              <small>Files or notes</small>
              <input type="file" onChange={(event) => void uploadFile(event.currentTarget.files?.[0])} />
            </label>
          </div>
        </GlassPanel>

        <GlassPanel className="focus-card-screen">
          <h3>Today's Focus <Sparkles size={15} /></h3>
          <small>Deep Work</small>
          <strong>25</strong>
          <span>min</span>
          <ActionButton onClick={() => void startFocusSessionTool({ subjectId: selected?.subjectId ?? snapshot.subjects[0]?.id ?? "art-9", minutes: 25 }).then(async (result) => {
            setNotice(result.summary);
            await refresh();
          })}>Start Focus</ActionButton>
          <div className="focus-orb-stage"><Orb size="large" /></div>
        </GlassPanel>

        <GlassPanel className="study-vault-card">
          <h3>Study Vault <Sparkles size={15} /></h3>
            {[
              ["Notes", snapshot.materials.length],
              ["Flashcards", snapshot.flashcards.length],
              ["Summaries", snapshot.highlights.length],
              ["Files", snapshot.materials.filter((material) => material.type !== "text").length],
            ].map(([label, count]) => (
            <div className="vault-stat-row" key={label as string}><FileText size={15} /><span>{label as string}</span><strong>{count as number}</strong></div>
          ))}
          <ActionButton tone="ghost" onClick={() => setQuery("")}>Open Study Vault</ActionButton>
        </GlassPanel>

        <GlassPanel className="subject-folder">
          <div className="panel-title"><h3>Subject Folder</h3><span>{visibleMaterials.length}</span></div>
          <div className="folder-list compact-folder">
            {visibleMaterials.slice(0, 7).map((material) => (
              <button key={material.id} className={selected?.id === material.id ? "selected" : ""} onClick={() => setSelectedId(material.id)} type="button">
                <FileText size={15} />
                <span>{material.title}</span>
              </button>
            ))}
            {visibleMaterials.length === 0 && <p className="empty-state-note">No study materials yet.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="reader-panel universal-reader">
          <div className="panel-title">
            <div><small>File / Note Reader</small><h3>{selected?.title ?? "No material selected"}</h3></div>
            <div className="mini-icon-row">
              <button onClick={() => setQuery(selected?.title ?? "")} disabled={!selected} title="Find this material" type="button"><Search size={16} /></button>
              <button onClick={() => void saveHighlight()} disabled={!selected} title="Save highlight" type="button"><Check size={16} /></button>
              <button onClick={() => void shareSelected()} disabled={!selected || snapshot.groups.length === 0} title={snapshot.groups.length === 0 ? "Create a group before sharing" : "Share to group"} type="button"><Share2 size={16} /></button>
            </div>
          </div>
          {notice && <p className="inline-notice">{notice}</p>}
          <article className="reader-content">{selected?.content ?? "Upload a file or note to start building your Study Vault."}</article>
          {selected && (
            <div className="material-tools">
              <button onClick={() => void renameSelected()} type="button">Rename</button>
              <input value={folderInput} onChange={(event) => setFolderInput(event.currentTarget.value)} placeholder={selected.folder || "Folder"} aria-label="Material folder" />
              <button onClick={() => void moveSelected()} type="button">Move</button>
              <input value={tagInput} onChange={(event) => setTagInput(event.currentTarget.value)} placeholder="Add tag" aria-label="Material tag" />
              <button onClick={() => void addTag()} type="button">Tag</button>
              <button onClick={() => void deleteSelected()} type="button">Delete</button>
            </div>
          )}
          <div className="button-row">
            <ActionButton tone="ghost" onClick={() => void saveHighlight()}><Check size={15} /> Save Highlight</ActionButton>
            <ActionButton onClick={() => void generateFlashcards()}><Sparkles size={15} /> Create Cards</ActionButton>
          </div>
        </GlassPanel>

        <GlassPanel className="summary-panel universal-summary">
          <h3>AI Summary <Sparkles size={15} /></h3>
          <p>{summary || "Add study material first, then generate summaries, flashcards, and quizzes from your own notes."}</p>
          <ActionButton onClick={() => selected ? void new FallbackProvider().summarize(selected.content).then(setSummary) : setSummary("Upload or create a note first.")}>Generate Summary <Sparkles size={15} /></ActionButton>
        </GlassPanel>

        <GlassPanel className="flashcard-panel mini-study-panel">
          <h3>Flashcards Hub</h3>
          <span>{snapshot.flashcards.length} cards</span>
          {activeCard && (
            <div className="tiny-card-preview">
              <small>{answerShown ? "A" : "Q"}</small>
              <p>{answerShown ? activeCard.back : activeCard.front}</p>
              <em>{activeCard.mastery}% mastery · due {formatDue(activeCard.dueAt)}</em>
            </div>
          )}
          <ActionButton tone="ghost" onClick={() => void generateFlashcards()}>Review Now</ActionButton>
        </GlassPanel>
        <GlassPanel className="mini-study-panel flash-study-mode">
          <h3>Flashcard Study Mode</h3>
          <div className="mini-stat-grid">
            <Metric label="New" value={snapshot.flashcards.filter((card) => !snapshot.flashcardReviews.some((review) => review.flashcardId === card.id)).length} />
            <Metric label="Review" value={snapshot.flashcardReviews.length} />
            <Metric label="Mastered" value={snapshot.flashcards.filter((card) => card.mastery >= 80).length} />
          </div>
          <button className="study-reveal-btn" onClick={() => setAnswerShown((shown) => !shown)} disabled={!activeCard} type="button">
            {answerShown ? "Hide answer" : "Reveal answer"}
          </button>
          <div className="button-row">
            {activeCard && (["again", "hard", "good", "easy"] as const).map((quality) => (
              <button key={quality} onClick={() => void reviewCard(quality)} type="button">{quality}</button>
            ))}
          </div>
        </GlassPanel>
        <GlassPanel className="mini-study-panel quiz-generator-card">
          <h3>Quiz Generator</h3>
          <p>Create custom quizzes from notes, flashcards, or files.</p>
          <ActionButton onClick={() => void generateQuiz()}>Generate Quiz</ActionButton>
        </GlassPanel>
        <GlassPanel className="quiz-panel mini-study-panel">
          <h3>Quiz Results / Weak Topics</h3>
          {activeQuiz && (
            <div className="quiz-session">
              <strong>{activeQuiz.title}</strong>
              {activeQuiz.questions.map((question, index) => (
                <label key={question.id}>
                  <span>{index + 1}. {question.prompt}</span>
                  {question.options?.length ? (
                    <select value={quizAnswers[question.id] ?? ""} onChange={(event) => setQuizAnswers({ ...quizAnswers, [question.id]: event.currentTarget.value })}>
                      <option value="">Choose an answer</option>
                      {question.options.map((option) => <option key={option} value={option}>{option}</option>)}
                    </select>
                  ) : (
                    <input value={quizAnswers[question.id] ?? ""} onChange={(event) => setQuizAnswers({ ...quizAnswers, [question.id]: event.currentTarget.value })} placeholder="Write your answer" />
                  )}
                </label>
              ))}
              <ActionButton onClick={() => void submitQuiz()}>Submit Quiz</ActionButton>
            </div>
          )}
          {activeQuizAttempt && <div className="quiz-score-ring"><strong>{activeQuizAttempt.score}%</strong><small>{activeQuizAttempt.weakTopics.join(", ") || "No weak topics saved"}</small></div>}
          {!activeQuiz && <p className="empty-state-note">No quizzes generated yet.</p>}
        </GlassPanel>
      </div>
    </div>
  );
}

function CalendarScreen({ snapshot, refresh, focusTarget }: { snapshot: LumaSnapshot; refresh: () => Promise<LumaSnapshot>; focusTarget?: AgentFocusTarget | null }) {
  const [view, setView] = useState<CalendarView>("Week");
  const focusedTask =
    focusTarget?.sourceType === "task" && focusTarget.sourceId
      ? snapshot.tasks.find((task) => task.id === focusTarget.sourceId)
      : undefined;
  const [date, setDate] = useState(focusedTask ? parseISO(focusedTask.dueAt) : new Date());
  const [searchOpen, setSearchOpen] = useState(false);
  const [addOpen, setAddOpen] = useState(false);
  const [newTaskTitle, setNewTaskTitle] = useState("");
  const [selectedTaskId, setSelectedTaskId] = useState(focusedTask?.id ?? "");
  const [taskDraft, setTaskDraft] = useState<TaskItem | null>(focusedTask ? { ...focusedTask } : null);
  const [newChecklistTitle, setNewChecklistTitle] = useState("");
  const [taskNotice, setTaskNotice] = useState("");

  const weekStart = startOfWeek(date, { weekStartsOn: 1 });
  const weekDays = weekdayLabels.map((label, index) => {
    const day = addDays(weekStart, index);
    return { label, date: format(day, "d"), active: isSameDay(day, date), weekday: index + 1 };
  });
  const hours = ["7 AM", "8 AM", "9 AM", "10 AM", "11 AM", "12 PM", "1 PM", "2 PM", "3 PM", "4 PM"];
  const events = snapshot.sessions.map((session) => {
    const subject = subjectById(snapshot.subjects, session.subjectId);
    return {
      day: session.weekday - 1,
      start: decimalHour(session.start),
      end: decimalHour(session.end),
      time: sessionTimeLabel(session),
      title: subject?.name ?? "Class",
      meta: `${subject?.code ?? ""}${subject?.code ? " · " : ""}Room ${session.room}`,
      tone: sessionTone(session),
      detail: `${session.teacher} · ${sessionModeLabel(session)}`,
      session,
    };
  });
  const selectedWeekday = ((date.getDay() + 6) % 7) + 1;
  const selectedDaySessions = snapshot.sessions
    .filter((session) => session.weekday === selectedWeekday)
    .sort((a, b) => minutesFromTime(a.start) - minutesFromTime(b.start));
  const weeklyPeriods = snapshot.sessions.reduce((total, session) => {
    if (!session.periodLabel) return total + 1;
    const match = session.periodLabel.match(/P(\d)(?:-P?(\d))?/);
    if (!match) return total + 1;
    const start = Number(match[1]);
    const end = Number(match[2] ?? match[1]);
    return total + Math.max(1, end - start + 1);
  }, 0);

  const positionFor = (start: number, end: number) => ({
    top: `${((start - 7.75) / 8.5) * 100}%`,
    height: `${Math.max(((end - start) / 8.5) * 100, 7.8)}%`,
  });

  const submitTask = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!newTaskTitle.trim()) return;
    await createTaskTool({
      title: newTaskTitle.trim(),
      subjectId: snapshot.subjects[0]?.id ?? "art-9",
      dueAt: new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59).toISOString(),
      priority: "medium",
      type: "homework",
      notes: "Created from LUMA Calendar.",
    });
    await refresh();
    setNewTaskTitle("");
    setAddOpen(false);
  };

  const dueRows = snapshot.tasks
    .filter((task) => task.status !== "done")
    .sort((a, b) => parseISO(a.dueAt).getTime() - parseISO(b.dueAt).getTime())
    .slice(0, 5);
  const selectedTask = snapshot.tasks.find((task) => task.id === selectedTaskId);
  const selectedChecklist = selectedTask ? snapshot.checklistItems.filter((item) => item.taskId === selectedTask.id) : [];

  const openTask = (task: TaskItem) => {
    setSelectedTaskId(task.id);
    setTaskDraft({ ...task });
    setNewChecklistTitle("");
    setTaskNotice("");
  };

  const saveTaskDraft = async () => {
    if (!taskDraft) return;
    const result = await updateTaskTool({
      taskId: taskDraft.id,
      patch: {
        title: taskDraft.title,
        subjectId: taskDraft.subjectId,
        dueAt: taskDraft.dueAt,
        priority: taskDraft.priority,
        type: taskDraft.type,
        notes: taskDraft.notes,
        estimatedMinutes: taskDraft.estimatedMinutes,
      },
    });
    setTaskNotice(result.summary);
    await refresh();
  };

  const setTaskStatus = async (status: TaskItem["status"]) => {
    if (!selectedTask) return;
    const result = await updateTaskStatusTool({ taskId: selectedTask.id, status });
    setTaskNotice(result.summary);
    await refresh();
  };

  const addChecklistItem = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!selectedTask || !newChecklistTitle.trim()) return;
    const result = await createChecklistItemTool({ taskId: selectedTask.id, title: newChecklistTitle });
    setNewChecklistTitle("");
    setTaskNotice(result.summary);
    await refresh();
  };

  const toggleChecklistItem = async (item: ChecklistItem) => {
    const result = await updateChecklistItemTool({ itemId: item.id, patch: { done: !item.done } });
    setTaskNotice(result.summary);
    await refresh();
  };

  const deleteSelectedTask = async () => {
    if (!selectedTask || !confirm(`Delete "${selectedTask.title}"? This can be undone from action history.`)) return;
    const result = await deleteTaskTool(selectedTask.id);
    setTaskNotice(result.summary);
    setSelectedTaskId("");
    setTaskDraft(null);
    await refresh();
  };

  return (
    <div className="luma-calendar-screen">
      <div className="calendar-user-actions" aria-label="LUMA status">
        <button type="button" aria-label="Notifications"><Bell size={22} /></button>
        <span className="home-avatar" aria-label="Win profile">W<i aria-hidden="true" /></span>
      </div>

      <header className="luma-calendar-topbar">
        <div className="calendar-title-block">
          <h1>Calendar <Sparkles size={25} /></h1>
          <p>{format(weekStart, "MMM d")} - {format(addDays(weekStart, 6), "MMM d, yyyy")}</p>
        </div>
        <div className="calendar-view-tabs" aria-label="Calendar views">
          {(["Day", "Week", "Month", "Exam Season"] as CalendarView[]).map((item) => (
            <button key={item} className={view === item ? "active" : ""} onClick={() => setView(item)} type="button">
              {item}
            </button>
          ))}
        </div>
        <div className="calendar-top-actions">
          <button onClick={() => setDate(addDays(date, -7))} aria-label="Previous week" type="button"><ChevronLeft size={18} /></button>
          <button onClick={() => setDate(new Date())} type="button">Today</button>
          <button onClick={() => setDate(addDays(date, 7))} aria-label="Next week" type="button"><ChevronRight size={18} /></button>
          <button onClick={() => setSearchOpen(!searchOpen)} aria-label="Search calendar" type="button"><Search size={21} /></button>
          <button className="calendar-add-glow" onClick={() => setAddOpen(true)} aria-label="Add item" type="button"><Plus size={22} /></button>
        </div>
      </header>

      {searchOpen && (
        <motion.div className="calendar-search-popover" initial={{ opacity: 0, y: -8 }} animate={{ opacity: 1, y: 0 }}>
          <Search size={16} />
          <input placeholder="Search classes, tasks, rooms, or events..." />
        </motion.div>
      )}

      <main className="luma-calendar-content">
        <section className="calendar-week-glass" aria-label="Week calendar">
          <div className="week-days-row">
            <div />
            {weekDays.map((day) => (
              <div key={day.label} className={day.active ? "active" : ""}>
                <span>{day.label}</span>
                <strong>{day.date}</strong>
              </div>
            ))}
          </div>
          <div className="week-grid-shell">
            <div className="time-axis-labels">
              {hours.map((hour) => <span key={hour}>{hour}</span>)}
            </div>
            <div className="week-lanes">
              <div className="now-line" style={{ top: `${((decimalHour(format(new Date(), "HH:mm")) - 7.75) / 8.5) * 100}%` }}><span>{format(new Date(), "HH:mm")}</span></div>
              {weekDays.map((day, dayIndex) => (
                <div className="week-lane" key={day.label}>
                  {events.filter((item) => item.day === dayIndex).map((eventItem) => (
                    <article
                      key={`${eventItem.day}-${eventItem.title}-${eventItem.start}`}
                      className={`calendar-event-chip ${eventItem.tone}`}
                      style={positionFor(eventItem.start, eventItem.end)}
                    >
                      <small>{eventItem.time}</small>
                      <strong>{eventItem.title}</strong>
                      {eventItem.meta && <span>{eventItem.meta}</span>}
                      {eventItem.detail && <em>{eventItem.detail}</em>}
                    </article>
                  ))}
                  {events.filter((item) => item.day === dayIndex).length === 0 && (
                    <div className="calendar-empty-block">
                      <strong>{day.weekday === 1 ? "No normal classes" : "Unassigned"}</strong>
                      <small>{day.weekday === 1 ? "Grey timetable row" : "No subject block"}</small>
                    </div>
                  )}
                </div>
              ))}
            </div>
          </div>
        </section>

        <aside className="calendar-side-stack">
          <GlassPanel className="calendar-side-card smart-card">
            <div className="side-card-title"><h3>Smart Timetable <Sparkles size={15} /></h3><span>{format(date, "EEE, MMM d")}</span></div>
            {selectedDaySessions.map((session) => {
              const subject = subjectById(snapshot.subjects, session.subjectId);
              return (
              <button className={`smart-row ${sessionTone(session)}`} key={session.id} type="button">
                <span />
                <strong>{subject?.name ?? "Class"}<small>{sessionTimeLabel(session)}</small></strong>
                <em>{subject?.code ?? ""}</em>
                <i>Room {session.room}</i>
              </button>
              );
            })}
            {selectedDaySessions.length === 0 && <p className="empty-state-note">No normal subject classes are scheduled for this day.</p>}
            <button className="side-link" disabled title="The full schedule is already visible in the week grid" type="button">View full schedule <ChevronRight size={14} /></button>
          </GlassPanel>

          <GlassPanel className="calendar-side-card due-card">
            <div className="side-card-title"><h3>Due Soon <Sparkles size={15} /></h3></div>
            {dueRows.map((task) => (
              <button className="due-side-row" key={task.id} onClick={() => openTask(task)} type="button">
                <FileText size={15} />
                <strong>{task.title}</strong>
                <span>{formatDue(task.dueAt)}</span>
              </button>
            ))}
            {dueRows.length === 0 && <p className="empty-state-note">No tasks yet. Add homework, exams, or project deadlines when assigned.</p>}
            <button className="side-link" disabled={snapshot.tasks.length === 0} onClick={() => snapshot.tasks[0] && openTask(snapshot.tasks[0])} title={snapshot.tasks.length === 0 ? "Add a task first" : "Open first task"} type="button">View all tasks <ChevronRight size={14} /></button>
          </GlassPanel>

          <GlassPanel className="calendar-side-card exam-card">
            <div className="side-card-title"><h3>Weekly Load <Sparkles size={15} /></h3></div>
            <div className="exam-card-body"><span className="exam-glyph"><Sparkles size={24} /></span><div><strong>{weeklyPeriods} scheduled periods</strong><small>Tuesday-Friday normal timetable</small><small>4 flipped · 2 self-study</small></div><b>{snapshot.subjects.length}<small>subjects</small></b></div>
            <button className="study-plan-btn" disabled title="Study plan generation needs the revision-plan tool layer first" type="button">Open Study Plan <ChevronRight size={14} /></button>
          </GlassPanel>

          <GlassPanel className="calendar-side-card focus-card">
            <div className="side-card-title"><h3>Focus Suggestion <Sparkles size={15} /></h3></div>
            <p>Thursday afternoon to Friday morning is the pressure point: Additional Mathematics repeats across both days.</p>
            <div className="focus-card-action"><button onClick={() => void startFocusSessionTool({ subjectId: snapshot.subjects[0]?.id ?? "art-9", minutes: 25 }).then(refresh)} type="button">Start Focus</button><span>25 min</span><Orb size="small" /></div>
          </GlassPanel>
        </aside>
      </main>

      {addOpen && (
        <div className="calendar-modal-overlay">
          <motion.form className="calendar-modal-container glass-panel quick-add-modal" onSubmit={submitTask} initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }}>
            <div className="modal-header-row"><h3>Add to Calendar</h3><button aria-label="Close" onClick={() => setAddOpen(false)} type="button"><X size={16} /></button></div>
            <input value={newTaskTitle} onChange={(event) => setNewTaskTitle(event.currentTarget.value)} placeholder="Add homework, exam, or project..." autoFocus />
            <div className="modal-action-buttons"><button className="modal-cancel-btn" onClick={() => setAddOpen(false)} type="button">Cancel</button><button className="modal-submit-btn" type="submit">Add</button></div>
          </motion.form>
        </div>
      )}

      {selectedTask && taskDraft && (
        <div className="calendar-modal-overlay">
          <motion.div className="calendar-modal-container detail-view glass-panel" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }}>
            <div className="modal-header-row">
              <div>
                <h3>Task Detail</h3>
                <p>{taskNotice || "Edit, complete, reopen, or break this task into steps."}</p>
              </div>
              <button aria-label="Close task detail" onClick={() => setSelectedTaskId("")} type="button"><X size={16} /></button>
            </div>

            <div className="modal-form-fields">
              <label>
                Title
                <input value={taskDraft.title} onChange={(event) => setTaskDraft({ ...taskDraft, title: event.currentTarget.value })} />
              </label>
              <div className="modal-form-row">
                <label>
                  Subject
                  <select value={taskDraft.subjectId} onChange={(event) => setTaskDraft({ ...taskDraft, subjectId: event.currentTarget.value })}>
                    {snapshot.subjects.map((subject) => <option key={subject.id} value={subject.id}>{subject.name}</option>)}
                  </select>
                </label>
                <label>
                  Priority
                  <select value={taskDraft.priority} onChange={(event) => setTaskDraft({ ...taskDraft, priority: event.currentTarget.value as TaskItem["priority"] })}>
                    <option value="low">Low</option>
                    <option value="medium">Medium</option>
                    <option value="high">High</option>
                  </select>
                </label>
              </div>
              <div className="modal-form-row">
                <label>
                  Type
                  <select value={taskDraft.type} onChange={(event) => setTaskDraft({ ...taskDraft, type: event.currentTarget.value as TaskItem["type"] })}>
                    <option value="homework">Homework</option>
                    <option value="test">Test</option>
                    <option value="presentation">Presentation</option>
                    <option value="project">Project</option>
                    <option value="revision">Revision</option>
                    <option value="admin">Admin</option>
                  </select>
                </label>
                <label>
                  Due
                  <input
                    type="datetime-local"
                    value={format(parseISO(taskDraft.dueAt), "yyyy-MM-dd'T'HH:mm")}
                    onChange={(event) => setTaskDraft({ ...taskDraft, dueAt: new Date(event.currentTarget.value).toISOString() })}
                  />
                </label>
              </div>
              <label>
                Notes
                <textarea value={taskDraft.notes} onChange={(event) => setTaskDraft({ ...taskDraft, notes: event.currentTarget.value })} />
              </label>
            </div>

            <section className="detail-modal-body">
              <h4>Checklist</h4>
              <div className="checklist">
                {selectedChecklist.map((item) => (
                  <label key={item.id}>
                    <input checked={item.done} onChange={() => void toggleChecklistItem(item)} type="checkbox" />
                    <span>{item.title}</span>
                  </label>
                ))}
                {selectedChecklist.length === 0 && <p className="empty-state-note">No checklist items yet.</p>}
              </div>
              <form className="chat-compose" onSubmit={addChecklistItem}>
                <input value={newChecklistTitle} onChange={(event) => setNewChecklistTitle(event.currentTarget.value)} placeholder="Add a step..." />
                <button type="submit" title="Add checklist item"><Plus size={15} /></button>
              </form>
            </section>

            <div className="detail-modal-action-buttons">
              <button className="modal-submit-btn" onClick={() => void saveTaskDraft()} type="button">Save Task</button>
              <button onClick={() => void setTaskStatus(selectedTask.status === "done" ? "todo" : "done")} type="button">
                {selectedTask.status === "done" ? "Reopen" : "Complete"}
              </button>
              <button onClick={() => void setTaskStatus("in-progress")} disabled={selectedTask.status === "in-progress"} type="button">Mark In Progress</button>
              <button className="modal-cancel-btn" onClick={() => void deleteSelectedTask()} type="button">Delete</button>
            </div>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function TogetherScreen({
  snapshot,
  refresh,
  focusTarget,
}: {
  snapshot: LumaSnapshot;
  refresh: () => Promise<LumaSnapshot>;
  focusTarget?: AgentFocusTarget | null;
}) {
  const group = snapshot.groups[0];
  const [message, setMessage] = useState("");
  const [groupName, setGroupName] = useState("Study Workspace");
  const initialTab = focusTarget?.sourceType === "file" ? "Files" : focusTarget?.sourceType === "task" ? "Tasks" : "Overview";
  const [activeMobileTab, setActiveMobileTab] = useState(initialTab);
  const [togetherNotice, setTogetherNotice] = useState("");

  const groupTasks = group ? snapshot.tasks.filter((task) => task.groupId === group.id) : [];
  const sharedFiles = group ? snapshot.materials.filter((material) => material.sharedWith.includes(group.id)) : [];
  const chatMessages = group ? snapshot.chats.filter((chat) => chat.groupId === group.id).sort((a, b) => parseISO(a.createdAt).getTime() - parseISO(b.createdAt).getTime()) : [];
  const completedGroupTasks = groupTasks.filter((task) => task.status === "done").length;
  const groupProgress = groupTasks.length ? Math.round((completedGroupTasks / groupTasks.length) * 100) : 0;
  const memberAvatars = group?.members.map((member) => member.avatar) ?? [];
  const board = [
    { title: "To Do", status: "todo", tone: "todo", tasks: groupTasks.filter((task) => task.status === "todo") },
    { title: "In Progress", status: "in-progress", tone: "progress", tasks: groupTasks.filter((task) => task.status === "in-progress") },
    { title: "Done", status: "done", tone: "done", tasks: groupTasks.filter((task) => task.status === "done") },
  ] as const;

  const avatarClass = (avatar: string, extra = "") =>
    `avatar avatar-${avatar.replace("+", "plus").toLowerCase().replace(/[^a-z0-9]/g, "")} ${avatar.startsWith("+") ? "avatar-count" : "avatar-photo"} ${extra}`.trim();
  const renderAvatar = (avatar: string, extra = "") => (
    <span className={avatarClass(avatar, extra)}>
      <span className="avatar-label">{avatar}</span>
    </span>
  );

  const sendMessage = async () => {
    if (!group || !message.trim()) return;
    const result = await postGroupMessageTool({ groupId: group.id, message });
    setTogetherNotice(result.summary);
    setMessage("");
    await refresh();
  };

  const shareLatest = async () => {
    if (!group) return;
    const material = snapshot.materials.find((item) => !item.sharedWith.includes(group.id));
    if (!material) return;
    if (!confirm(`Share "${material.title}" with ${group.name}?`)) return;
    const result = await shareMaterialToGroupTool({ materialId: material.id, groupId: group.id });
    setTogetherNotice(result.summary);
    await refresh();
  };

  const addGroupTask = async () => {
    if (!group) return;
    const result = await createTaskTool({
      title: "Summarize project progress",
      subjectId: group.subjectId,
      dueAt: addDays(new Date(), 2).toISOString(),
      priority: "medium",
      type: "project",
      notes: "Generated from group workspace.",
      groupId: group.id,
      assignee: "You",
    });
    setTogetherNotice(result.summary);
    await refresh();
  };

  const editChat = async (chat: ChatMessage) => {
    const next = prompt("Edit message", chat.message)?.trim();
    if (!next || next === chat.message) return;
    const result = await updateGroupMessageTool({ chatId: chat.id, message: next });
    setTogetherNotice(result.success ? result.summary : result.error?.message ?? "Message was not updated.");
    await refresh();
  };

  const deleteChat = async (chat: ChatMessage) => {
    if (!confirm("Delete this message from group chat?")) return;
    const result = await deleteGroupMessageTool(chat.id);
    setTogetherNotice(result.success ? result.summary : result.error?.message ?? "Message was not deleted.");
    await refresh();
  };

  const setGroupTaskStatus = async (task: TaskItem, status: TaskItem["status"]) => {
    const result = await updateTaskStatusTool({ taskId: task.id, status });
    setTogetherNotice(result.success ? result.summary : result.error?.message ?? "Task was not updated.");
    await refresh();
  };

  const removeSharedFile = async (material: Material) => {
    if (!group || !confirm(`Remove "${material.title}" from ${group.name}?`)) return;
    const result = await unshareMaterialFromGroupTool({ materialId: material.id, groupId: group.id });
    setTogetherNotice(result.success ? result.summary : result.error?.message ?? "Shared file was not removed.");
    await refresh();
  };

  const changeMyRole = async (role: string) => {
    if (!group || !confirm(`Change your group role to ${role}?`)) return;
    const result = await updateGroupMemberRoleTool({ groupId: group.id, memberId: snapshot.profile.id, role });
    setTogetherNotice(result.success ? result.summary : result.error?.message ?? "Role was not updated.");
    await refresh();
  };

  const createGroup = async (event: React.FormEvent) => {
    event.preventDefault();
    const result = await createGroupTool({
      name: groupName,
      subjectId: snapshot.subjects[0]?.id ?? "art-9",
    });
    if (result.success) {
      setGroupName("Study Workspace");
      await refresh();
    }
  };

  if (!group) {
    return (
      <div className="together-workspace first-use">
        <header className="together-header">
          <div>
            <h1>Together <Sparkles size={18} /></h1>
            <p>Create a private workspace before LUMA shows shared files, messages, or group AI.</p>
          </div>
          <div className="together-actions">
            <button type="button" title="Notifications"><Bell size={18} /></button>
            <span className="profile-dot avatar-photo avatar-profile"><span className="avatar-label">{snapshot.profile.avatar}</span><i /></span>
          </div>
        </header>

        <GlassPanel className="group-summary together-card">
          <div className="group-summary-top">
            <div>
              <h2>Start a Group Workspace <Sparkles size={18} /></h2>
              <p>Group data stays empty until you deliberately create or share it.</p>
              <small>Private by default · group AI uses shared workspace content only</small>
            </div>
            <span className="status-pill beta">First use</span>
          </div>
          <form className="project-title-field" onSubmit={createGroup}>
            <span>Workspace name</span>
            <input value={groupName} onChange={(event) => setGroupName(event.currentTarget.value)} placeholder="Biology project team" />
            <button className="modal-submit-btn" type="submit">Create Workspace</button>
          </form>
        </GlassPanel>
      </div>
    );
  }

  return (
    <div className={`together-workspace tab-${activeMobileTab.toLowerCase()}`}>
      <header className="together-header">
        <div>
          <h1>Together <Sparkles size={18} /></h1>
          <p>Study together. Build together. Grow together.</p>
        </div>
        <div className="together-actions">
          <button type="button" title="Notifications"><Bell size={18} /></button>
          <span className="profile-dot avatar-photo avatar-profile"><span className="avatar-label">{snapshot.profile.avatar}</span><i /></span>
        </div>
      </header>

      <nav className="together-tabs" aria-label="Together sections">
        {["Friends", "Group Overview", "Shared Files", "Group Chat"].map((item) => (
          <button key={item} className={item === "Group Overview" ? "active" : ""} type="button">{item}</button>
        ))}
      </nav>

      <div className="together-mobile-tabs">
        {["Overview", "Tasks", "Files", "Chat", "AI"].map((item) => (
          <button key={item} className={activeMobileTab === item ? "active" : ""} onClick={() => setActiveMobileTab(item)} type="button">{item}</button>
        ))}
      </div>

      {togetherNotice && <p className="empty-state-note together-status" role="status">{togetherNotice}</p>}

      <section className="together-grid">
        <GlassPanel className="friends-panel together-card">
          <div className="together-card-head">
            <h3>Friends <Sparkles size={14} /></h3>
            <button disabled title="Friend search needs sync to be enabled first" type="button"><Plus size={15} /></button>
          </div>
          <label className="mini-search">
            <Search size={14} />
            <input placeholder="Search friends" />
          </label>
          <div className="friend-list">
            {group.members.map((member) => (
              <button className="friend-row" key={member.id} type="button">
                {renderAvatar(member.avatar)}
                <span><strong>{member.name}</strong><small>{member.role}</small></span>
                <i className={member.online ? "online" : "offline"} />
              </button>
            ))}
            {group.members.length === 0 && <p className="empty-state-note">No members yet.</p>}
          </div>
          <label className="role-select-row">
            <span>Your role</span>
            <select value={group.members.find((member) => member.id === snapshot.profile.id)?.role ?? "owner"} onChange={(event) => void changeMyRole(event.currentTarget.value)}>
              <option value="owner">owner</option>
              <option value="admin">admin</option>
              <option value="member">member</option>
              <option value="viewer">viewer</option>
            </select>
          </label>
          <button className="quiet-link" disabled title="All members are already shown" type="button">View all members <ChevronRight size={14} /></button>
        </GlassPanel>

        <GlassPanel className="study-streak-panel together-card">
          <div className="together-card-head">
            <h3>Study Streak <Sparkles size={14} /></h3>
          </div>
          <div className="streak-summary">
            <span className="streak-flame"><Flame size={20} /></span>
            <strong>{snapshot.profile.streakDays}</strong>
            <small>days</small>
          </div>
          <div className="streak-bars" aria-label="Weekly study streak">
            {[34, 54, 64, 78, 92, 100, 96].map((height, index) => (
              <span key={index} style={{ "--bar-height": `${height}%` } as CSSProperties}>
                <i />
                <small>{["M", "T", "W", "T", "F", "S", "S"][index]}</small>
              </span>
            ))}
          </div>
          <p>Keep it up.</p>
        </GlassPanel>

        <GlassPanel className="group-summary together-card">
          <div className="group-summary-top">
            <div>
              <h2>{group.name} <Sparkles size={18} /></h2>
              <p>{subjectById(snapshot.subjects, group.subjectId)?.name ?? "Shared study workspace"}</p>
              <small>{group.members.length} member{group.members.length === 1 ? "" : "s"} · {groupTasks.length} group task{groupTasks.length === 1 ? "" : "s"}</small>
            </div>
            <span className="status-pill">Active</span>
          </div>
          <div className="member-stack">
            {memberAvatars.map((avatar) => <Fragment key={avatar}>{renderAvatar(avatar)}</Fragment>)}
            <button disabled title="Invites require synced accounts" type="button">Invite</button>
          </div>
          <div className="summary-body">
            <div className="progress-ring" style={{ "--progress": `${groupProgress}%` } as CSSProperties}>
              <strong>{groupProgress}%</strong>
              <span>Complete</span>
            </div>
            <div className="summary-stats">
              <span><Clock size={18} /><strong>{Math.round(snapshot.studySessions.filter((session) => session.mode === "group").reduce((sum, session) => sum + session.minutes, 0) / 60 * 10) / 10}h</strong><small>Group Focus</small></span>
              <span><Check size={18} /><strong>{completedGroupTasks} / {groupTasks.length}</strong><small>Tasks Done</small></span>
            </div>
          </div>
          <div className="session-strip">
            <span><CalendarDays size={20} /><strong>{groupTasks[0]?.dueAt ? formatDue(groupTasks[0].dueAt) : "No deadline"}</strong><small>Next group task</small></span>
            <span><Gauge size={20} /><strong>{sharedFiles.length}</strong><small>Shared files</small></span>
            <span><Trophy size={20} /><strong>{chatMessages.length}</strong><small>Messages</small></span>
          </div>
        </GlassPanel>

        <GlassPanel className="shared-files together-card">
          <div className="together-card-head">
            <h3>Shared Files <Sparkles size={14} /></h3>
            <button onClick={() => void shareLatest()} disabled={snapshot.materials.every((material) => material.sharedWith.includes(group.id))} title={snapshot.materials.length === 0 ? "Upload a Study Vault material first" : "Share latest unshared material"} type="button">Share latest <ChevronRight size={14} /></button>
          </div>
          <div className="file-list">
            {sharedFiles.map((material) => (
              <article className="file-row" key={material.id}>
                <span className={`file-type ${material.type}`}><FileText size={15} /></span>
                <span><strong>{material.title}</strong><small>{material.type.toUpperCase()} · Updated {formatDue(material.updatedAt)}</small></span>
                <button onClick={() => void removeSharedFile(material)} type="button">Remove</button>
              </article>
            ))}
            {sharedFiles.length === 0 && <p className="empty-state-note">No shared files yet. Share a Study Vault material when the group needs it.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="chat-panel together-card">
          <div className="together-card-head">
            <h3>Group Chat <Sparkles size={14} /></h3>
            <button disabled title="Chat actions are available after selecting a saved message" type="button">...</button>
          </div>
          <div className="compact-chat">
            {chatMessages.map((chat) => (
              <article key={chat.id}>
                {renderAvatar(chat.author.split(" ").map((part) => part[0]).join(""))}
                <p><strong>{chat.author}</strong> <small>{format(parseISO(chat.createdAt), "p")}</small><br />{chat.message}</p>
                {chat.author === "You" && (
                  <div className="chat-message-actions">
                    <button onClick={() => void editChat(chat)} type="button">Edit</button>
                    <button onClick={() => void deleteChat(chat)} type="button">Delete</button>
                  </div>
                )}
              </article>
            ))}
            {chatMessages.length === 0 && <p className="empty-state-note">No messages yet. Start with a useful update.</p>}
          </div>
          <form
            className="chat-compose"
            onSubmit={(event) => {
              event.preventDefault();
              void sendMessage();
            }}
          >
            <input value={message} onChange={(event) => setMessage(event.currentTarget.value)} placeholder="Message group..." />
            <button type="submit" title="Send message"><Send size={16} /></button>
          </form>
        </GlassPanel>

        <GlassPanel className="task-board-panel together-card">
          <div className="together-card-head">
            <h3>Group Task Board <Sparkles size={14} /></h3>
            <button onClick={() => void addGroupTask()} type="button"><Plus size={14} /> Add task</button>
          </div>
          <div className="together-kanban">
            {board.map((column) => (
              <section className={`together-column ${column.tone}`} key={column.title}>
                <header><strong>{column.title}</strong><span>{column.tasks.length}</span></header>
                {column.tasks.map((task) => (
                  <article key={task.id}>
                    <strong>{task.title}</strong>
                    <div><small>{task.priority}</small><small>{formatDue(task.dueAt)}</small>{task.status === "done" ? <Check size={14} /> : renderAvatar(task.assignee ?? "You", "mini")}</div>
                    <div className="task-move-actions">
                      {(["todo", "in-progress", "done"] as const).filter((status) => status !== task.status).map((status) => (
                        <button key={status} onClick={() => void setGroupTaskStatus(task, status)} type="button">{status === "in-progress" ? "Start" : status === "done" ? "Done" : "Reopen"}</button>
                      ))}
                    </div>
                  </article>
                ))}
                {column.tasks.length === 0 && <p className="empty-state-note">No {column.title.toLowerCase()} tasks.</p>}
                <button onClick={() => void addGroupTask()} type="button"><Plus size={13} /> Add task</button>
              </section>
            ))}
          </div>
        </GlassPanel>

        <GlassPanel className="project-timeline together-card">
          <div className="together-card-head">
            <h3>Group Project Timeline <Sparkles size={14} /></h3>
            <div className="timeline-toggle"><button className="active" type="button">Week</button><button type="button">Month</button></div>
          </div>
          <div className="timeline-board" aria-label="Group project timeline">
            <div className="timeline-axis">{["Task", "Now", "+1d", "+2d", "+3d", "+4d", "+5d", "+6d"].map((item) => <span key={item}>{item}</span>)}</div>
            {groupTasks.slice(0, 5).map((task, index) => (
              <div className="timeline-row" key={task.id}>
                <strong>{task.title}</strong>
                <span className={`timeline-bar start-${Math.min(index + 1, 7)} span-2`}><b>{formatDue(task.dueAt)}</b>{renderAvatar(task.assignee ?? "You", "mini")}</span>
              </div>
            ))}
            {groupTasks.length === 0 && <p className="empty-state-note">Add a group task to build a timeline.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="shared-ai together-card">
          <div className="together-card-head">
            <h3>Shared AI Context / LUMA Assistant <Sparkles size={14} /></h3>
            <span className="status-pill beta">BETA</span>
          </div>
          <div className="ai-context-grid">
            <section><strong>Group Context</strong><p>{group.name}<br />Shared files: {sharedFiles.length}<br />Open tasks: {groupTasks.filter((task) => task.status !== "done").length}</p><div className="member-stack small">{memberAvatars.slice(0, 4).map((avatar) => <Fragment key={avatar}>{renderAvatar(avatar)}</Fragment>)}</div></section>
            <section><strong>Recent Topics</strong><div className="topic-cloud">{sharedFiles.flatMap((material) => material.tags).slice(0, 5).map((topic) => <span key={topic}>{topic}</span>)}{sharedFiles.length === 0 && <span>No shared file tags yet</span>}</div></section>
            <section><strong>Ask LUMA for the group</strong><button disabled={sharedFiles.length === 0} onClick={() => queueGlobalAgentPrompt("Summarize our shared group files and cite the sources.", true)} title={sharedFiles.length === 0 ? "Share files before group AI can use them" : "Ask LUMA to summarize shared files"} type="button">Summarize shared files</button><button disabled={groupTasks.length === 0} onClick={() => queueGlobalAgentPrompt("What group tasks are unfinished and who owns them?", true)} title={groupTasks.length === 0 ? "Add tasks before LUMA can summarize progress" : "Ask LUMA to find unfinished tasks"} type="button">Find unfinished tasks</button><button disabled={!group || (sharedFiles.length === 0 && groupTasks.length === 0)} onClick={() => queueGlobalAgentPrompt("Create a meeting agenda from our group progress, shared files, and unfinished tasks.", true)} title={!group ? "Create a group first" : "Ask LUMA to create a meeting agenda"} type="button">Create meeting agenda</button></section>
            <section><strong>Smart Suggestions</strong><p>{groupTasks.length ? `You have ${groupTasks.filter((task) => task.status !== "done").length} unfinished group task${groupTasks.filter((task) => task.status !== "done").length === 1 ? "" : "s"}.` : "Create a first group task to unlock useful progress suggestions."}</p><button disabled={groupTasks.length === 0} title="Add group tasks first" type="button">View recommendations <ChevronRight size={13} /></button></section>
          </div>
        </GlassPanel>
      </section>

    </div>
  );
}

function Metric({ label, value, metricKey }: { label: string; value: string | number; metricKey?: ProfileMetricKey }) {
  const unit =
    metricKey === "streak"
      ? "days"
      : metricKey === "focusHours"
        ? "hrs"
        : metricKey === "classesTracked"
          ? "classes"
          : metricKey === "tasksDone"
            ? "tasks"
            : "";
  const icon =
    metricKey === "streak"
      ? "flame"
      : metricKey === "focusHours"
        ? "clock"
        : metricKey === "classesTracked"
          ? "book"
          : metricKey === "tasksDone"
            ? "check"
            : "spark";
  return (
    <div className={`metric metric-${icon}`}>
      <span>{label}</span>
      <div>
        <i aria-hidden="true" />
        <strong>{value}</strong>
      </div>
      {unit && <em>{unit}</em>}
    </div>
  );
}

function CreateScreen({
  snapshot,
  refresh,
  focusTarget,
}: {
  snapshot: LumaSnapshot;
  refresh: () => Promise<LumaSnapshot>;
  focusTarget?: AgentFocusTarget | null;
}) {
  const createTemplates: {
    type: DocumentType;
    description: string;
    tint: string;
    icon: LucideIcon;
  }[] = [
    { type: "Essay", description: "Build a structured argument.", tint: "violet", icon: FileText },
    { type: "Report", description: "Research and write clearly.", tint: "blue", icon: BookOpen },
    { type: "Presentation", description: "Turn ideas into a visual story.", tint: "rose", icon: PanelTop },
    { type: "Reflection", description: "Capture learning and growth.", tint: "coral", icon: Sparkles },
    { type: "Study Guide", description: "Make revision material from notes.", tint: "mint", icon: BookOpen },
    { type: "Project Plan", description: "Turn work into milestones and tasks.", tint: "cyan", icon: LayoutGrid },
  ];
  const createSteps = [
    { label: "Idea", detail: "Define topic", icon: Wand2 },
    { label: "Outline", detail: "Structure content", icon: FileText },
    { label: "Write", detail: "Draft & refine", icon: FilePlus2 },
    { label: "Review", detail: "Polish & perfect", icon: Check },
    { label: "Export", detail: "Share & submit", icon: Share2 },
  ];
  const detailFields = [
    ["Subject", snapshot.subjects[0]?.name ?? ""],
    ["Teacher", snapshot.subjects[0]?.teacher ?? ""],
    ["Audience", "M.5"],
    ["Tone", "Academic"],
    ["Collaboration", "Solo draft"],
    ["Linked Group", ""],
  ];
  const contextSources = [
    { label: "Teacher instructions", icon: FileText },
    { label: "Rubric", icon: Gauge },
    { label: "Study vault notes", icon: HardDrive },
    { label: "Previous draft", icon: RotateCcw },
    { label: "Links", icon: Globe2 },
    { label: "Group files", icon: Users },
  ];
  const [type, setType] = useState<DocumentType>("Essay");
  const [title, setTitle] = useState("Untitled Project");
  const focusedDocumentId =
    focusTarget?.sourceType === "document" && focusTarget.sourceId && snapshot.documents.some((document) => document.id === focusTarget.sourceId)
      ? focusTarget.sourceId
      : "";
  const [selectedId, setSelectedId] = useState(focusedDocumentId || snapshot.documents[0]?.id || "");
  const [projectNotice, setProjectNotice] = useState("");
  const [newOutlineItem, setNewOutlineItem] = useState("");
  const selected =
    snapshot.documents.find((document) => document.id === selectedId) ??
    snapshot.documents[0] ?? {
      id: "doc-draft-preview",
      type,
      title,
      outline: ["Introduction", "Main argument", "Evidence", "Review", "Conclusion"],
      body: "Start with the assignment prompt, paste notes, or ask LUMA to shape a first outline.",
      status: "draft",
      dueAt: addDays(new Date(), 7).toISOString(),
      sourceMaterialIds: [],
      milestones: ["Choose Type", "Project Details", "Add Context", "Build Outline", "Finish"],
    };
  const selectedProgress = selected.status === "ready" ? 100 : selected.status === "review" ? 86 : 72;
  const selectedVersions = snapshot.documentVersions
    .filter((version) => version.documentId === selected.id)
    .sort((a, b) => parseISO(b.createdAt).getTime() - parseISO(a.createdAt).getTime())
    .slice(0, 5);

  const createOutline = async () => {
    const outline = await new FallbackProvider().createOutline(title, type, snapshot.materials.map((material) => material.content));
    const result = await createDocumentTool({
      type,
      title,
      outline,
      body: outline.map((item) => `${item}\n`).join("\n"),
      dueAt: addDays(new Date(), 7).toISOString(),
      sourceMaterialIds: snapshot.materials.slice(0, 2).map((material) => material.id),
      milestones: ["Details", "Sources", "Outline", "Generate"],
    });
    if (result.data) {
      setSelectedId(result.data.id);
    }
    setProjectNotice(result.summary);
    await refresh();
  };

  const updateBody = async (body: string) => {
    if (!snapshot.documents.some((document) => document.id === selected.id)) return;
    const result = await updateDocumentTool({
      documentId: selected.id,
      patch: { body },
      summary: `Saved "${selected.title}".`,
      versionLabel: "Autosave",
    });
    setProjectNotice(result.summary);
    await refresh();
  };

  const renameDocument = async (document: CreateDocument) => {
    const nextTitle = prompt("Rename project", document.title)?.trim();
    if (!nextTitle || nextTitle === document.title) return;
    const result = await updateDocumentTool({
      documentId: document.id,
      patch: { title: nextTitle },
      summary: `Renamed "${document.title}" to "${nextTitle}".`,
      versionLabel: "Before rename",
    });
    setTitle(nextTitle);
    setProjectNotice(result.summary);
    await refresh();
  };

  const duplicateDocument = async (document: CreateDocument) => {
    const result = await duplicateDocumentTool(document.id);
    if (result.data) setSelectedId(result.data.id);
    setProjectNotice(result.summary);
    await refresh();
  };

  const archiveDocument = async (document: CreateDocument) => {
    const result = await updateDocumentTool({
      documentId: document.id,
      patch: { status: "archived" },
      summary: `Archived "${document.title}".`,
      versionLabel: "Before archive",
    });
    setProjectNotice(result.summary);
    await refresh();
  };

  const deleteDocument = async (document: CreateDocument) => {
    if (!confirm(`Delete "${document.title}"? This can be undone from the local action history.`)) return;
    const result = await deleteDocumentTool(document.id);
    setProjectNotice(result.summary);
    if (selectedId === document.id) setSelectedId("");
    await refresh();
  };

  const improveWriting = async (action: string) => {
    if (!snapshot.documents.some((document) => document.id === selected.id)) {
      setProjectNotice("Create a project before using writing actions.");
      return;
    }
    const trimmed = selected.body.trim();
    const nextBody =
      action === "Shorten"
        ? trimmed.split(/\s+/).slice(0, Math.max(40, Math.floor(trimmed.split(/\s+/).length * 0.7))).join(" ")
        : action === "Expand Content"
          ? `${selected.body.trim()}\n\nAdd one concrete example, one source-backed detail, and one reflection before submitting.`
          : action === "Check Grammar"
            ? selected.body.replace(/\s+/g, " ").trim()
            : action === "Make It Academic"
              ? selected.body.replace(/\bI think\b/gi, "The evidence suggests").replace(/\bkids\b/gi, "students")
              : action === "Enhance Structure"
                ? `Claim\n${selected.body.trim()}\n\nEvidence\nAdd a source or example here.\n\nReflection\nExplain why this matters.`
                : `${selected.body.trim()}\n\nRevision note: clarify the claim, add evidence, and keep the conclusion concise.`;
    const result = await updateDocumentTool({
      documentId: selected.id,
      patch: { body: nextBody },
      summary: `${action} applied to "${selected.title}".`,
      versionLabel: `Before ${action}`,
    });
    setProjectNotice(result.summary);
    await refresh();
  };

  const updateOutlineItem = async (index: number, value: string) => {
    if (!snapshot.documents.some((document) => document.id === selected.id)) return;
    const outline = selected.outline.map((item, itemIndex) => (itemIndex === index ? value : item)).filter((item) => item.trim());
    const result = await updateDocumentTool({
      documentId: selected.id,
      patch: { outline },
      summary: `Updated outline for "${selected.title}".`,
      versionLabel: "Before outline edit",
    });
    setProjectNotice(result.summary);
    await refresh();
  };

  const addOutlineItem = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!snapshot.documents.some((document) => document.id === selected.id) || !newOutlineItem.trim()) return;
    const result = await updateDocumentTool({
      documentId: selected.id,
      patch: { outline: [...selected.outline, newOutlineItem.trim()] },
      summary: `Added outline section to "${selected.title}".`,
      versionLabel: "Before adding outline section",
    });
    setNewOutlineItem("");
    setProjectNotice(result.summary);
    await refresh();
  };

  const restoreVersion = async (versionId: string) => {
    if (!confirm("Restore this version? LUMA will save the current draft as a version first.")) return;
    const result = await restoreDocumentVersionTool(versionId);
    setProjectNotice(result.success ? result.summary : result.error?.message ?? "Version was not restored.");
    await refresh();
  };

  const exportDocx = async (document = selected) => {
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({ children: [new TextRun({ text: document.title, bold: true, size: 32 })] }),
            ...document.body.split("\n").map((line) => new Paragraph(line)),
          ],
        },
      ],
    });
    downloadBlob(await Packer.toBlob(doc), `${document.title}.docx`);
  };

  const exportPptx = async () => {
    const deck = new pptxgen();
    selected.outline.forEach((line) => {
      const slide = deck.addSlide();
      slide.background = { color: "151b42" };
      slide.addText(line, { x: 0.6, y: 0.7, w: 8, h: 0.6, color: "FFFFFF", fontSize: 28, bold: true });
      slide.addText(selected.title, { x: 0.6, y: 1.5, w: 8, h: 1, color: "B9B3FF", fontSize: 16 });
    });
    await deck.writeFile({ fileName: `${selected.title}.pptx` });
  };

  const exportHtml = () => {
    const blob = new Blob([`<h1>${selected.title}</h1><pre>${selected.body}</pre>`], { type: "text/html" });
    downloadBlob(blob, `${selected.title}.html`);
  };

  return (
    <div className="universal-screen create-universal">
      <UniversalHeader title="Create" subtitle="From ideas to impact." avatar={snapshot.profile.avatar} />
      <div className="create-dashboard">
        <GlassPanel className="create-hub">
          <div className="create-panel-heading">
            <div>
              <h3>Create Hub <Sparkles size={15} /></h3>
              <p>Start something new or continue what matters.</p>
            </div>
            <ActionButton onClick={() => void createOutline()}><Plus size={15} /> New project</ActionButton>
          </div>
          <div className="create-types">
            {createTemplates.map((item) => {
              const Icon = item.icon;
              return (
                <button
                  key={item.type}
                  className={type === item.type ? "selected" : ""}
                  data-tint={item.tint}
                  onClick={() => setType(item.type)}
                  type="button"
                >
                  <span><Icon size={18} /></span>
                  <strong>{item.type}</strong>
                  <small>{item.description}</small>
                </button>
              );
            })}
          </div>
        </GlassPanel>

        <GlassPanel className="creation-flow">
          <h3>Document Creation Flow <Sparkles size={15} /></h3>
          <div className="flow-track">
            {createSteps.map((step, index) => {
              const Icon = step.icon;
              return (
                <span key={step.label} className={index === 2 ? "active" : ""}>
                  <i><Icon size={16} /></i>
                  <b>{index + 1} {step.label}</b>
                  <small>{step.detail}</small>
                </span>
              );
            })}
          </div>
        </GlassPanel>

        <GlassPanel className="recent-projects-panel">
          <div className="create-panel-heading">
            <div>
              <h3>Recent Projects <Clock size={15} /></h3>
              <p>Open, export, or clean up current work.</p>
            </div>
          </div>
          {projectNotice && <p className="empty-state-note" role="status">{projectNotice}</p>}
          <div className="recent-project-list">
            {snapshot.documents.length === 0 && <p className="empty-state-note">No projects yet. Choose a type, set a title, and create your first outline.</p>}
            {[...snapshot.documents]
              .filter((document) => document.status !== "archived" || document.id === selectedId)
              .sort((a, b) => (a.id === selected.id ? -1 : b.id === selected.id ? 1 : 0))
              .slice(0, 3)
              .map((document, index) => (
              <article key={`${document.id}-${index}`} className={document.id === selected.id ? "selected" : ""}>
                <button type="button" onClick={() => setSelectedId(document.id)}>
                  <strong>{document.title}</strong>
                  <small>{document.type} · {document.status}</small>
                  <span className="progress-line"><b style={{ width: `${document.id === selected.id ? selectedProgress : 54}%` }} /></span>
                  <em>Due {document.dueAt ? formatDue(document.dueAt) : "No due date"}</em>
                </button>
                <div className="project-actions">
                  <button type="button" title="Open" onClick={() => setSelectedId(document.id)}><Circle size={10} /><span>Open</span></button>
                  <button type="button" title="Rename" onClick={() => void renameDocument(document)}><Circle size={10} /><span>Rename</span></button>
                  <button type="button" title="Duplicate" onClick={() => void duplicateDocument(document)}><Circle size={10} /><span>Duplicate</span></button>
                  <button disabled type="button" title="Folders are not wired yet"><Circle size={10} /><span>Move to folder</span></button>
                  <button type="button" title="Export" onClick={() => void exportDocx(document)}><Download size={13} /><span>Export</span></button>
                  <button disabled type="button" title="Share links require sync to be enabled"><Share2 size={13} /><span>Share</span></button>
                  <button type="button" title="Archive" onClick={() => void archiveDocument(document)}><Circle size={10} /><span>Archive</span></button>
                  <button type="button" title="Delete" onClick={() => void deleteDocument(document)}><Trash2 size={13} /><span>Delete</span></button>
                </div>
              </article>
            ))}
          </div>
        </GlassPanel>

        <GlassPanel className="project-flow-panel">
          <h3>New Project Flow <Sparkles size={15} /></h3>
          <div className="project-step-row">
            {["Choose Type", "Project Details", "Add Context"].map((step, index) => (
              <span key={step} className={index === 1 ? "active" : ""}>{index + 1}. {step}</span>
            ))}
          </div>
          <label className="project-title-field">
            <span>Title</span>
            <input value={title} onChange={(event) => setTitle(event.currentTarget.value)} />
          </label>
          <div className="project-detail-grid">
            {detailFields.map(([label, value]) => (
              <label key={label}>
                <span>{label}</span>
                <input defaultValue={value} />
              </label>
            ))}
            <label>
              <span>Due Date</span>
              <input type="datetime-local" defaultValue="2026-05-28T23:59" />
            </label>
            <label>
              <span>Estimated Length</span>
              <select defaultValue="1200-1500 words"><option>1200-1500 words</option><option>5-7 slides</option><option>Brief outline</option></select>
            </label>
          </div>
          <div className="context-source-grid">
            {contextSources.map((source) => {
              const Icon = source.icon;
              return <button key={source.label} type="button"><Icon size={14} />{source.label}</button>;
            })}
          </div>
        </GlassPanel>

        <GlassPanel className="assignment-planner">
          <h3>Presentation / Assignment Planner <Sparkles size={15} /></h3>
          <label>
            Type
            <select value={type} onChange={(event) => setType(event.currentTarget.value as DocumentType)}>
              {documentTypes.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>
          </label>
          <label>
            Audience
            <select><option>High School</option><option>College</option><option>Team</option></select>
          </label>
          <label>
            Length
            <select><option>1200-1500 words</option><option>5-7 slides</option><option>Brief outline</option></select>
          </label>
          <label>
            Due Date
            <input type="datetime-local" defaultValue="2026-05-28T23:59" />
          </label>
          <div className="smart-reminder"><Bell size={16} /><span>Smart Reminder <small>1 day before due date</small></span></div>
        </GlassPanel>

        <GlassPanel className="editor-panel structured-editor">
          <div className="editor-toolbar">
            <select aria-label="Block style" defaultValue="Heading 2"><option>Heading 2</option><option>Body</option><option>Quote</option></select>
            {["B", "I", "U"].map((tool) => <button disabled key={tool} title={`${tool} formatting is not wired yet`} type="button">{tool}</button>)}
            <button disabled type="button" title="Bulleted list formatting is not wired yet"><LayoutGrid size={14} /></button>
            <button disabled type="button" title="Links are not wired yet"><Globe2 size={14} /></button>
            <button disabled type="button" title="Quote blocks are not wired yet"><Info size={14} /></button>
            <span><Check size={13} /> Saved</span>
          </div>
          <div className="document-page">
            <label className="title-inline">
              <span>Title</span>
              <input value={title} onChange={(event) => setTitle(event.currentTarget.value)} />
            </label>
            <textarea value={selected.body} onChange={(event) => void updateBody(event.currentTarget.value)} />
          </div>
          <div className="editor-status">
            <span>{selected.body.trim().split(/\s+/).filter(Boolean).length} words</span>
            <span>{Math.max(1, Math.ceil(selected.body.trim().split(/\s+/).filter(Boolean).length / 220))} min read</span>
            <span>English (US)</span>
            <span>{selectedVersions.length} versions</span>
          </div>
        </GlassPanel>

        <GlassPanel className="assistant-panel">
          <h3><Orb size="small" /> LUMA AI Assistant <Sparkles size={15} /></h3>
          <p>How can I help with your document?</p>
          <div className="assistant-action-group">
            <strong>Writing Actions</strong>
            {["Improve Writing", "Expand Content", "Shorten", "Check Grammar", "Enhance Structure", "Make It Academic"].map((action) => (
              <button key={action} onClick={() => void improveWriting(action)} type="button"><Wand2 size={14} />{action}</button>
            ))}
          </div>
          <ActionButton disabled title="Citation search needs source-grounding work before it can be enabled"><Sparkles size={15} /> Suggest Citations</ActionButton>
          <div className="assistant-tuning">
            <label>Tone<select defaultValue="Academic"><option>Academic</option><option>Clear</option><option>Reflective</option></select></label>
            <label>Creativity<select defaultValue="Balanced"><option>Balanced</option><option>Conservative</option><option>Expressive</option></select></label>
          </div>
        </GlassPanel>

        <GlassPanel className="outline-panel">
          <div className="panel-title">
            <h3>AI Outline Generation <Sparkles size={15} /></h3>
            <button type="button" onClick={() => void createOutline()}><RefreshCw size={13} /> Regenerate</button>
          </div>
          {selected.outline.map((item, index) => (
            <div className="outline-row" key={`${item}-${index}`}>
              <span>{["I.", "II.", "III.", "IV.", "V.", "VI."][index] ?? `${index + 1}.`}</span>
              <input value={item} onChange={(event) => void updateOutlineItem(index, event.currentTarget.value)} aria-label={`Outline section ${index + 1}`} />
              <button onClick={() => void updateOutlineItem(index, "")} type="button" title="Delete outline section"><Trash2 size={13} /></button>
            </div>
          ))}
          <form className="outline-add-row" onSubmit={addOutlineItem}>
            <input value={newOutlineItem} onChange={(event) => setNewOutlineItem(event.currentTarget.value)} placeholder="Add outline section" />
            <button type="submit"><Plus size={13} /> Add</button>
          </form>
        </GlassPanel>

        <GlassPanel className="version-panel">
          <div className="panel-title">
            <h3>Version History <RotateCcw size={15} /></h3>
            <span>{selectedVersions.length}</span>
          </div>
          {selectedVersions.map((version) => (
            <button key={version.id} onClick={() => void restoreVersion(version.id)} type="button">
              <strong>{version.label}</strong>
              <small>{format(parseISO(version.createdAt), "MMM d, p")} · {version.body.trim().split(/\s+/).filter(Boolean).length} words</small>
            </button>
          ))}
          {selectedVersions.length === 0 && <p className="empty-state-note">Autosaved versions appear after your first edit.</p>}
        </GlassPanel>

        <GlassPanel className="export-panel">
          <h3>Export / Share <Sparkles size={15} /></h3>
          <div className="export-grid">
            <button onClick={() => window.print()} type="button">
              <Download size={18} />
              PDF
            </button>
            <button onClick={() => void exportDocx()} type="button">
              <Download size={18} />
              Word
            </button>
            <button onClick={exportHtml} type="button">
              <Download size={18} />
              HTML
            </button>
            <button onClick={() => void exportPptx()} type="button">
              <Download size={18} />
              PPTX
            </button>
          </div>
          <div className="share-link-row">
            <span>Share Link <small>Requires sync before external sharing</small></span>
            <button disabled title="Share links require sync to be enabled" type="button">Copy Link</button>
          </div>
        </GlassPanel>

        <GlassPanel className="notes-document-panel">
          <div className="tab-row"><span className="active">Notes</span><span>Transcript</span><span>Upload</span></div>
          <ul>
            {["Social media connects youth globally", "Can improve communication & collaboration", "Risk of anxiety, comparison, cyberbullying", "Affects sleep and attention span", "Importance of digital balance", "Need for education and awareness"].map((note) => <li key={note}>{note}</li>)}
          </ul>
          <ActionButton onClick={() => void createOutline()}><Sparkles size={15} /> Convert to Document</ActionButton>
        </GlassPanel>
      </div>
    </div>
  );
}

function ProfileScreen({
  snapshot,
  saveProfile,
  navigate,
}: {
  snapshot: LumaSnapshot;
  saveProfile: (profile: UserProfile) => Promise<void>;
  navigate: (route: RouteId, target?: Omit<AgentFocusTarget, "route">) => void;
}) {
  const [analyticsRange, setAnalyticsRange] = useState<"week" | "month">("week");
  const [activeInsightId, setActiveInsightId] = useState<string | null>(null);
  const [dismissedInsights, setDismissedInsights] = useState<string[]>([]);
  const [editOpen, setEditOpen] = useState(false);
  const [showAllAchievements, setShowAllAchievements] = useState(false);
  const [focusInput, setFocusInput] = useState("");
  const [draftProfile, setDraftProfile] = useState<UserProfile>(() => buildProfileDraft(snapshot.profile));
  const [selectedSubjectId, setSelectedSubjectId] = useState(snapshot.profile.currentGoal?.linkedSubjectId ?? snapshot.subjects[0]?.id ?? "");
  const [selectedDayId, setSelectedDayId] = useState<string | null>(null);
  const [subjectInspectOpen, setSubjectInspectOpen] = useState(false);

  const blankGoal = (): StudyGoal => ({
    id: uid("goal"),
    title: "",
    type: "exam",
    progress: 0,
    dueAt: format(addDays(new Date(), 7), "yyyy-MM-dd'T'HH:mm"),
    targetFocusMinutes: 300,
    linkedSubjectId: snapshot.subjects[0]?.id,
    linkedTaskIds: [],
    linkedMaterialIds: [],
    status: "active",
  });

  const openEditProfile = () => {
    setDraftProfile({
      ...buildProfileDraft(snapshot.profile),
      currentGoal: snapshot.profile.currentGoal ? { ...snapshot.profile.currentGoal } : blankGoal(),
    });
    setFocusInput("");
    setEditOpen(true);
  };

  const queueAgentPrompt = (prompt: string) => {
    queueGlobalAgentPrompt(prompt);
  };

  const periodStart = analyticsRange === "week" ? startOfWeek(new Date(), { weekStartsOn: 1 }) : startOfMonth(new Date());
  const periodEnd = analyticsRange === "week" ? endOfWeek(new Date(), { weekStartsOn: 1 }) : endOfMonth(new Date());
  const chartDays = eachDayOfInterval({ start: periodStart, end: periodEnd });
  const periodSessions = snapshot.studySessions.filter((session) => {
    const date = parseISO(session.startedAt);
    return session.completed && date >= periodStart && date <= periodEnd;
  });
  const periodTasksDone = snapshot.tasks.filter((task) => task.status === "done" && parseISO(task.dueAt) >= periodStart && parseISO(task.dueAt) <= periodEnd);
  const previousRangeEnd = addDays(periodStart, -1);
  const previousRangeStart = addDays(previousRangeEnd, -(chartDays.length - 1));
  const previousSessions = snapshot.studySessions.filter((session) => {
    const date = parseISO(session.startedAt);
    return session.completed && date >= previousRangeStart && date <= previousRangeEnd;
  });

  const focusByDay = chartDays.map((day) => {
    const sessions = periodSessions.filter((session) => isSameDay(parseISO(session.startedAt), day));
    const minutes = sessions.reduce((sum, session) => sum + session.minutes, 0);
    const completedTasks = periodTasksDone.filter((task) => isSameDay(parseISO(task.dueAt), day)).length;
    return {
      id: format(day, "yyyy-MM-dd"),
      label: analyticsRange === "week" ? format(day, "EEE") : format(day, "MMM d"),
      minutes,
      hours: Number((minutes / 60).toFixed(1)),
      completedTasks,
      sessions,
    };
  });
  const previousTotalMinutes = previousSessions.reduce((sum, session) => sum + session.minutes, 0);
  const totalMinutes = periodSessions.reduce((sum, session) => sum + session.minutes, 0);
  const trendDelta = previousTotalMinutes === 0 ? 0 : Math.round(((totalMinutes - previousTotalMinutes) / previousTotalMinutes) * 100);
  const selectedDay = focusByDay.find((day) => day.id === selectedDayId) ?? focusByDay.find((day) => day.minutes > 0) ?? focusByDay[focusByDay.length - 1];

  const subjectMinutes = new Map<string, number>();
  periodSessions.forEach((session) => {
    subjectMinutes.set(session.subjectId, (subjectMinutes.get(session.subjectId) ?? 0) + session.minutes);
  });
  const subjectBreakdown = snapshot.subjects
    .map((subject) => ({
      subjectId: subject.id,
      name: subject.name,
      minutes: subjectMinutes.get(subject.id) ?? 0,
      color: subject.color,
    }))
    .filter((subject) => subject.minutes > 0)
    .sort((a, b) => b.minutes - a.minutes)
    .map((subject) => ({
      ...subject,
      value: totalMinutes === 0 ? 0 : Math.round((subject.minutes / totalMinutes) * 100),
    }));
  const selectedSubject = subjectById(snapshot.subjects, selectedSubjectId) ?? subjectById(snapshot.subjects, subjectBreakdown[0]?.subjectId ?? snapshot.subjects[0]?.id ?? "");
  let subjectDonutCursor = 0;
  const subjectDonutBackground = subjectBreakdown.length
    ? `conic-gradient(${subjectBreakdown
        .map((subject) => {
          const start = subjectDonutCursor;
          const portion = totalMinutes > 0 ? (subject.minutes / totalMinutes) * 100 : 0;
          subjectDonutCursor += portion;
          return `${subject.color} ${start}% ${subjectDonutCursor}%`;
        })
        .join(", ")}, ${subjectBreakdown[0]?.color ?? "#9f8cff"} ${subjectDonutCursor}% 100%)`
    : "conic-gradient(#9f8cff 0% 100%)";
  const selectedSubjectMetrics = selectedSubject
    ? {
        focusMinutes: subjectMinutes.get(selectedSubject.id) ?? 0,
        taskProgress: snapshot.tasks.filter((task) => task.subjectId === selectedSubject.id),
        materials: snapshot.materials.filter((material) => material.subjectId === selectedSubject.id).slice(0, 3),
        weakTopics: snapshot.quizzes.filter((quiz) => quiz.subjectId === selectedSubject.id).flatMap((quiz) => quiz.weakTopics).slice(0, 3),
        recentSessions: periodSessions.filter((session) => session.subjectId === selectedSubject.id).slice(-3).reverse(),
      }
    : null;

  const timeWindows = periodSessions.reduce<Record<string, number>>((accumulator, session) => {
    const hour = parseISO(session.startedAt).getHours();
    const key = hour < 12 ? "Morning" : hour < 17 ? "Afternoon" : hour < 22 ? "Evening" : "Late Night";
    accumulator[key] = (accumulator[key] ?? 0) + session.minutes;
    return accumulator;
  }, {});
  const bestWindow = Object.entries(timeWindows).sort((a, b) => b[1] - a[1])[0]?.[0] ?? "Evening";
  const mostProductiveDay = focusByDay.reduce((best, day) => (day.minutes > best.minutes ? day : best), focusByDay[0]);
  const leastStudiedSubject = [...subjectBreakdown].sort((a, b) => a.minutes - b.minutes)[0];
  const hasStudyActivity = periodSessions.length > 0 || periodTasksDone.length > 0;
  const completedTasks = snapshot.tasks.filter((task) => task.status === "done");
  const avgDailyFocus = focusByDay.length > 0 ? Math.round(totalMinutes / focusByDay.length) : 0;
  const upcomingDeadlines = tasksDueWithin(snapshot.tasks, 7);
  const metricValues: Record<ProfileMetricKey, string | number> = {
    streak: snapshot.profile.streakDays,
    focusHours: Math.round(snapshot.profile.focusMinutes / 60),
    classesTracked: `${snapshot.subjects.length}`,
    tasksDone: `${completedTasks.length}`,
    flashcardsReviewed: `${snapshot.flashcardReviews.length}`,
    quizzesCompleted: `${snapshot.quizAttempts.length}`,
    groupContributions: `${snapshot.tasks.filter((task) => task.groupId).length}`,
    averageDailyFocus: `${avgDailyFocus} min`,
    upcomingDeadlines: `${upcomingDeadlines.length}`,
  };
  const visibleMetrics = snapshot.profile.metricsVisible.slice(0, 4).map((metricKey) => {
    const definition = profileMetricDefinitions.find((metric) => metric.key === metricKey) ?? profileMetricDefinitions[0];
    return { ...definition, value: metricValues[metricKey] };
  });

  const currentGoal = snapshot.profile.currentGoal;
  const linkedGoalSubject = currentGoal?.linkedSubjectId ? subjectById(snapshot.subjects, currentGoal.linkedSubjectId) : undefined;
  const goalDueLabel = currentGoal?.dueAt
    ? (() => {
        const due = parseISO(currentGoal.dueAt);
        const dayCount = Math.max(0, differenceInCalendarDays(due, new Date()));
        return `Due in ${dayCount} days · ${format(due, "MMM d")}`;
      })()
    : "No deadline yet";
  const goalRemainingHours = currentGoal?.targetFocusMinutes
    ? Math.max(0, Math.round((currentGoal.targetFocusMinutes * (1 - currentGoal.progress / 100)) / 60))
    : null;

  const insights = hasStudyActivity ? [
    {
      id: "focus-day",
      text: `${mostProductiveDay?.label ?? "This week"} has been your strongest focus day in this period.`,
      evidence: `${Math.round((mostProductiveDay?.minutes ?? 0) / 60 * 10) / 10} hours across ${mostProductiveDay?.sessions.length ?? 0} sessions.`,
      action: "Build Study Plan",
      prompt: `Build a study plan around my strongest ${analyticsRange} focus window and current goal.`,
    },
    {
      id: "focus-window",
      text: `${bestWindow} sessions appear to support your deeper work best.`,
      evidence: `${Math.round((timeWindows[bestWindow] ?? 0) / 60 * 10) / 10} hours recorded in that window.`,
      action: "Show Evidence",
      prompt: `Summarize the focus evidence behind my best ${bestWindow.toLowerCase()} study window.`,
    },
    {
      id: "subject-balance",
      text: `${leastStudiedSubject?.name ?? "A subject"} has received less time in this period.`,
      evidence: `${Math.round(((leastStudiedSubject?.minutes ?? 0) / 60) * 10) / 10} hours versus ${subjectBreakdown[0]?.name ?? "your top subject"} at ${Math.round(((subjectBreakdown[0]?.minutes ?? 0) / 60) * 10) / 10} hours.`,
      action: "Build Study Plan",
      prompt: `Help rebalance my weekly study time with more support for ${leastStudiedSubject?.name ?? "my timetable"}.`,
    },
    {
      id: "task-completion",
      text: "Completed work is trending steadily, not just piling into one day.",
      evidence: `${periodTasksDone.length} completed items landed inside this ${analyticsRange} view.`,
      action: "Show Evidence",
      prompt: "Show me the evidence behind my recent task completion trend.",
    },
  ].filter((insight) => !dismissedInsights.includes(insight.id)) : [];
  const visibleInsights = insights.slice(0, 3);

  const timeline = [...snapshot.profile.journey]
    .sort((a, b) => parseISO(b.date).getTime() - parseISO(a.date).getTime())
    .slice(0, 6);
  const achievementItems = showAllAchievements ? snapshot.profile.achievements : snapshot.profile.achievements.slice(0, 6);
  const preferenceItems = [
    ...snapshot.profile.preferencesSnapshot,
    `Theme: ${snapshot.settings.theme}`,
    `${snapshot.settings.focusSessionMinutes}-minute Focus Sessions`,
    `${snapshot.settings.focusSound} Ambience`,
    snapshot.settings.timeFormat === "24-hour" ? "24-hour Time" : "12-hour Time",
    snapshot.settings.aiMemoryEnabled ? "Local AI Enabled" : "Local AI Off",
  ].filter((item, index, items) => items.indexOf(item) === index).slice(0, 5);

  const avatarText = initialsForName(snapshot.profile.name);
  const avatarStyle: CSSProperties =
    snapshot.profile.avatarStyle === "photo" && snapshot.profile.avatarUrl
      ? {
          backgroundImage: `url("${snapshot.profile.avatarUrl}")`,
          backgroundSize: `${snapshot.profile.avatarCrop?.zoom ?? 100}%`,
          backgroundPosition: `${snapshot.profile.avatarCrop?.x ?? 50}% ${snapshot.profile.avatarCrop?.y ?? 50}%`,
        }
      : {
          background: snapshot.profile.avatarGradient ?? profileAvatarGradients[0],
        };

  const draftAvatarText = initialsForName(draftProfile.name || snapshot.profile.name);
  const draftAvatarStyle: CSSProperties =
    draftProfile.avatarStyle === "photo" && draftProfile.avatarUrl
      ? {
          backgroundImage: `url("${draftProfile.avatarUrl}")`,
          backgroundSize: `${draftProfile.avatarCrop?.zoom ?? 100}%`,
          backgroundPosition: `${draftProfile.avatarCrop?.x ?? 50}% ${draftProfile.avatarCrop?.y ?? 50}%`,
        }
      : {
          background: draftProfile.avatarGradient ?? profileAvatarGradients[0],
        };

  const updateDraftGoal = (patch: Partial<StudyGoal>) => {
    setDraftProfile((current) => ({
      ...current,
      currentGoal: {
        ...(current.currentGoal ?? blankGoal()),
        ...patch,
      },
    }));
  };

  const addFocusArea = () => {
    const nextValue = focusInput.trim();
    if (!nextValue || draftProfile.focusSubjects.includes(nextValue) || draftProfile.focusSubjects.length >= 6) return;
    setDraftProfile((current) => ({ ...current, focusSubjects: [...current.focusSubjects, nextValue] }));
    setFocusInput("");
  };

  const toggleMetric = (key: ProfileMetricKey) => {
    setDraftProfile((current) => {
      const active = current.metricsVisible.includes(key);
      if (active) {
        return { ...current, metricsVisible: current.metricsVisible.filter((item) => item !== key) };
      }
      if (current.metricsVisible.length >= 4) return current;
      return { ...current, metricsVisible: [...current.metricsVisible, key] };
    });
  };

  const saveProfileDraft = async () => {
    const trimmedName = draftProfile.name.trim() || snapshot.profile.name;
    const nextProfile: UserProfile = {
      ...draftProfile,
      name: trimmedName,
      year: `${draftProfile.gradeOrYear || "Student"}${draftProfile.program ? ` · ${draftProfile.program}` : ""}`,
      avatar: initialsForName(trimmedName),
      bio: draftProfile.bio?.trim() || undefined,
      motto: draftProfile.motto?.trim() || draftProfile.bio?.trim() || snapshot.profile.motto,
      focusSubjects: draftProfile.focusSubjects.slice(0, 6),
      metricsVisible: draftProfile.metricsVisible.length > 0 ? draftProfile.metricsVisible.slice(0, 4) : snapshot.profile.metricsVisible,
      currentGoal:
        draftProfile.currentGoal && draftProfile.currentGoal.title.trim()
          ? {
              ...draftProfile.currentGoal,
              title: draftProfile.currentGoal.title.trim(),
              dueAt: draftProfile.currentGoal.dueAt || undefined,
              progress: clamp(draftProfile.currentGoal.progress, 0, 100),
            }
          : undefined,
    };
    await saveProfile(nextProfile);
    setEditOpen(false);
  };

  return (
    <div className="universal-screen profile-universal">
      <UniversalHeader title="Profile" subtitle="Personal. Intelligent. Private." avatar={snapshot.profile.avatar} />
      <div className="profile-dashboard">
        <GlassPanel className="profile-card identity-card">
          <div className="profile-identity-shell">
            <div className="profile-identity">
              <span className={`profile-photo ${snapshot.profile.avatarStyle === "photo" ? "has-image" : ""}`} style={avatarStyle}>
                {snapshot.profile.avatarStyle === "photo" ? "" : avatarText}
              </span>
              <div>
                <div className="profile-heading-row">
                  <div>
                    <h2>{snapshot.profile.name}</h2>
                    <p>{snapshot.profile.year}</p>
                  </div>
                  <button className="profile-edit-trigger" onClick={openEditProfile} type="button">
                    <FilePlus2 size={14} />
                    Edit Profile
                  </button>
                </div>
                <blockquote>"{snapshot.profile.motto || snapshot.profile.bio || "M.5 IM timetable ready."}"</blockquote>
                <div className="profile-meta">
                  {snapshot.profile.location && (
                    <span>
                      <MapPin size={14} />
                      {snapshot.profile.location}
                    </span>
                  )}
                  <span>
                    <Clock size={14} />
                    {formatProfileTime(snapshot.profile.timezone)}
                  </span>
                  {snapshot.profile.language && (
                    <span>
                      <Languages size={14} />
                      {snapshot.profile.language}
                    </span>
                  )}
                </div>
                <div className="profile-supporting-meta">
                  {snapshot.profile.school && <span>{snapshot.profile.school}</span>}
                  {snapshot.profile.gradeOrYear && <span>{snapshot.profile.gradeOrYear}</span>}
                  {snapshot.profile.program && <span>{snapshot.profile.program}</span>}
                </div>
                {snapshot.profile.bio && <p className="profile-bio">{snapshot.profile.bio}</p>}
              </div>
            </div>
            <div className="profile-privacy-row">
              <span>
                <Lock size={14} />
                Profile {snapshot.profile.visibility.profile}
              </span>
              <span>
                <Shield size={14} />
                Avatar {snapshot.profile.visibility.avatar}
              </span>
              <span>
                <Users size={14} />
                Focus {snapshot.profile.visibility.studyFocus}
              </span>
            </div>
          </div>
        </GlassPanel>

        <GlassPanel className="study-focus-card">
          <div className="panel-title">
            <h3>Study Focus <Sparkles size={15} /></h3>
            <button onClick={openEditProfile} type="button">Edit Focus</button>
          </div>
          <div className="chip-row profile-focus-tags">
            {snapshot.profile.focusSubjects.map((subject) => (
              <button key={subject} type="button" onClick={() => navigate("learn")}>
                {subject}
              </button>
            ))}
            {snapshot.profile.focusSubjects.length === 0 && <p className="empty-state-note">No pinned focus areas yet.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="goal-card">
          <div className="panel-title">
            <h3>Current Goal <Sparkles size={15} /></h3>
            <span className={`profile-status-pill ${currentGoal?.status === "complete" ? "complete" : currentGoal?.status === "archived" ? "archived" : "active"}`}>
              {currentGoal?.status === "complete" ? "Completed" : currentGoal?.status === "archived" ? "Archived" : "Active"}
            </span>
          </div>
          {currentGoal ? (
            <>
              <strong className="goal-title">{currentGoal.title}</strong>
              <div className="goal-detail-row">
                {linkedGoalSubject && <span>{linkedGoalSubject.name}</span>}
                {goalRemainingHours !== null && <span>{goalRemainingHours} hrs left</span>}
              </div>
              <div className="progress-line"><b style={{ width: `${currentGoal.progress}%` }} /></div>
              <div className="goal-progress-meta">
                <small>{goalDueLabel}</small>
                <em>{currentGoal.progress}%</em>
              </div>
              <div className="goal-action-row">
                <ActionButton onClick={() => navigate("learn")}><BookOpen size={15} /> Continue Plan</ActionButton>
                <ActionButton tone="ghost" onClick={openEditProfile}><FilePlus2 size={15} /> Edit Goal</ActionButton>
                <ActionButton tone="ghost" onClick={() => queueAgentPrompt(`Help me reach "${currentGoal.title}" before ${goalDueLabel}.`)}><Sparkles size={15} /> Ask LUMA</ActionButton>
              </div>
            </>
          ) : (
            <div className="goal-empty">
              <p>Create one meaningful goal so Profile can reflect what matters now.</p>
              <ActionButton onClick={openEditProfile}><Plus size={15} /> Create Goal</ActionButton>
            </div>
          )}
        </GlassPanel>

        <GlassPanel className="glance-card">
          <div className="panel-title">
            <h3>At a Glance <Sparkles size={15} /></h3>
            <button onClick={openEditProfile} type="button">Customize</button>
          </div>
          <div className="overview-metrics">
            {visibleMetrics.map((metric) => (
              <Metric key={metric.key} metricKey={metric.key} label={metric.label} value={metric.value} />
            ))}
          </div>
        </GlassPanel>

        <GlassPanel className="analytics-panel">
          <div className="panel-title">
            <h3>Study Analytics <Sparkles size={15} /></h3>
            <div className="profile-segmented">
              <button className={analyticsRange === "week" ? "active" : ""} onClick={() => setAnalyticsRange("week")} type="button">This Week</button>
              <button className={analyticsRange === "month" ? "active" : ""} onClick={() => setAnalyticsRange("month")} type="button">This Month</button>
            </div>
          </div>
          <div className="analytics-summary-row">
            <span>{Math.round(totalMinutes / 60 * 10) / 10} hrs total</span>
            <span>{trendDelta >= 0 ? "+" : ""}{trendDelta}% vs prior period</span>
            <span>{periodSessions.length} sessions logged</span>
          </div>
          <div className="charts-grid profile-charts-grid">
            <div className={`analytics-chart-card ${selectedDayId ? "is-inspecting" : ""}`}>
              <div className="chart-card-head">
                <h4>Focus Time</h4>
                <p>Click a day below to inspect sessions and completed work.</p>
              </div>
              <ResponsiveContainer width="100%" height={128}>
                <BarChart data={focusByDay}>
                  <CartesianGrid stroke="rgba(255,255,255,.08)" vertical={false} />
                  <XAxis dataKey="label" stroke="#a9adff" tickLine={false} axisLine={false} />
                  <YAxis stroke="#a9adff" tickLine={false} axisLine={false} />
                  <Tooltip contentStyle={chartTooltipStyle} />
                  <Bar dataKey="hours" fill="#9f8cff" radius={[10, 10, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
              <div className="focus-day-strip">
                {focusByDay.map((day) => (
                  <button
                    key={day.id}
                    className={selectedDay?.id === day.id ? "active" : ""}
                    onClick={() => setSelectedDayId(day.id)}
                    type="button"
                  >
                    <strong>{day.label}</strong>
                    <small>{day.hours}h</small>
                  </button>
                ))}
              </div>
              {selectedDayId && selectedDay && (
                <div className="focus-day-detail">
                  <strong>{selectedDay.label}</strong>
                  <span>{selectedDay.hours} hrs across {selectedDay.sessions.length} sessions</span>
                  <small>{selectedDay.completedTasks} tasks completed on this day</small>
                </div>
              )}
            </div>

            <div className={`analytics-chart-card ${subjectInspectOpen ? "is-inspecting-subject" : ""}`}>
              <div className="chart-card-head">
                <h4>Time by Subject</h4>
                <p>Open a subject to see materials, weak topics, and recent sessions.</p>
              </div>
              <button
                className="subject-donut"
                style={{ "--subject-donut": subjectDonutBackground } as CSSProperties}
                onClick={() => {
                  if (!selectedSubject) return;
                  setSelectedSubjectId(selectedSubject.id);
                  setSubjectInspectOpen(true);
                }}
                onDoubleClick={() => setSubjectInspectOpen((current) => !current)}
                type="button"
                aria-label={selectedSubject ? `Inspect ${selectedSubject.name}` : "Inspect subject focus"}
              >
                <span />
              </button>
              <div className="subject-list">
                {subjectBreakdown.map((subject) => (
                  <button
                    key={subject.subjectId}
                    className={selectedSubject?.id === subject.subjectId ? "active" : ""}
                    onClick={() => {
                      setSelectedSubjectId(subject.subjectId);
                      setSubjectInspectOpen(true);
                    }}
                    type="button"
                  >
                    <i style={{ background: subject.color }} aria-hidden="true" />
                    <span>{subject.name}</span>
                    <strong>{subject.value}%</strong>
                  </button>
                ))}
                {subjectBreakdown.length === 0 && <p className="empty-state-note">No focus time logged yet.</p>}
              </div>
              {subjectInspectOpen && selectedSubject && selectedSubjectMetrics && (
                <div className="subject-inspect-card">
                  <strong>{selectedSubject.name}</strong>
                  <small>{Math.round(selectedSubjectMetrics.focusMinutes / 60 * 10) / 10} hrs in this view</small>
                  <p>{selectedSubjectMetrics.materials.length} linked materials · {selectedSubjectMetrics.taskProgress.filter((task) => task.status === "done").length}/{selectedSubjectMetrics.taskProgress.length || 0} tasks done</p>
                  {selectedSubjectMetrics.weakTopics.length > 0 && <span>Weak topics: {selectedSubjectMetrics.weakTopics.join(", ")}</span>}
                  {selectedSubjectMetrics.recentSessions[0] && <span>Most recent session: {format(parseISO(selectedSubjectMetrics.recentSessions[0].startedAt), "MMM d, p")}</span>}
                </div>
              )}
            </div>

            <div className="analytics-chart-card">
              <div className="chart-card-head">
                <h4>Productivity Trend</h4>
                <p>Patterns are directional, not diagnostic.</p>
              </div>
              <ResponsiveContainer width="100%" height={94}>
                <AreaChart data={focusByDay}>
                  <CartesianGrid stroke="rgba(255,255,255,.05)" vertical={false} />
                  <XAxis dataKey="label" stroke="#a9adff" tickLine={false} axisLine={false} />
                  <YAxis hide />
                  <Tooltip contentStyle={chartTooltipStyle} />
                  <Area type="monotone" dataKey="completedTasks" stroke="#83dfc6" fill="rgba(131,223,198,.16)" />
                  <Area type="monotone" dataKey="hours" stroke="#b9a9ff" fill="rgba(159,140,255,.26)" />
                </AreaChart>
              </ResponsiveContainer>
              <div className="pattern-callout">
                <strong>Pattern</strong>
                <p>{hasStudyActivity ? `You tend to do your best focused work in the ${bestWindow.toLowerCase()}.` : "Start focus sessions to build a real activity pattern."}</p>
              </div>
            </div>

            <div className="key-insights">
              <div className="chart-card-head">
                <h4>Key Insights <Sparkles size={13} /></h4>
                <p>Every insight stays grounded in traceable activity.</p>
              </div>
              {visibleInsights.map((insight) => (
                <article key={insight.id} className={`insight-row ${activeInsightId === insight.id ? "active" : ""}`}>
                  <div className="insight-copy">
                    <p><Check size={14} /> {insight.text}</p>
                    {activeInsightId === insight.id && <small>{insight.evidence}</small>}
                  </div>
                  <div className="insight-actions">
                    <button onClick={() => queueAgentPrompt(insight.prompt)} type="button">{insight.action}</button>
                    <button onClick={() => setActiveInsightId((current) => (current === insight.id ? null : insight.id))} type="button">
                      {activeInsightId === insight.id ? "Hide Evidence" : "Show Evidence"}
                    </button>
                    <button onClick={() => setDismissedInsights((current) => [...current, insight.id])} type="button">Dismiss</button>
                    <button onClick={() => setDismissedInsights((current) => [...current, insight.id])} type="button">Do Not Suggest Again</button>
                  </div>
                </article>
              ))}
              {visibleInsights.length === 0 && <p className="empty-state-note">No insights yet. Add tasks or complete focus sessions to generate real patterns.</p>}
            </div>
          </div>
        </GlassPanel>

        <GlassPanel className="achievements-card">
          <div className="panel-title">
            <h3>Achievements <Sparkles size={15} /></h3>
            <button onClick={() => setShowAllAchievements((current) => !current)} type="button">{showAllAchievements ? "Show Less" : "View All"}</button>
          </div>
          <div className="badge-row badge-grid">
            {achievementItems.map((badge) => (
              <article key={badge.id} className="badge-card">
                <span>{badge.icon === "users" ? <Users size={18} /> : badge.icon === "clock" ? <Clock size={18} /> : badge.icon === "book" ? <BookOpen size={18} /> : <Sparkles size={18} />}</span>
                <strong>{badge.title}</strong>
                <small>{badge.description}</small>
                <em>{badge.unlockedAt ? `Unlocked ${format(parseISO(badge.unlockedAt), "MMM d, yyyy")}` : badge.progress && badge.maxProgress ? `${badge.progress}/${badge.maxProgress}` : "In progress"}</em>
              </article>
            ))}
            {achievementItems.length === 0 && <p className="empty-state-note">No achievements yet.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="journey-card">
          <div className="panel-title">
            <h3>Study Journey <Sparkles size={15} /></h3>
            <span>Private record of growth</span>
          </div>
          <div className="journey-list">
            {timeline.map((entry) => (
              <article key={entry.id} className="journey-item">
                <strong>{format(parseISO(entry.date), "MMM d")}</strong>
                <div>
                  <h4>{entry.title}</h4>
                  <p>{entry.detail}</p>
                </div>
              </article>
            ))}
            {timeline.length === 0 && <p className="empty-state-note">No study journey entries yet.</p>}
          </div>
        </GlassPanel>

        <GlassPanel className="preferences-card">
          <div className="panel-title">
            <h3>Quick Actions <Sparkles size={15} /></h3>
          </div>
          <div className="profile-actions-stack">
            <button onClick={openEditProfile} type="button"><FilePlus2 size={15} /><span><strong>Edit Profile</strong><small>Update your info and preferences</small></span><ChevronRight size={15} /></button>
            <button onClick={() => navigate("settings")} type="button"><BarChart3 size={15} /><span><strong>View Progress</strong><small>Detailed stats and reports</small></span><ChevronRight size={15} /></button>
            <button onClick={() => queueAgentPrompt("Help me decide what to work on next based on my profile.")} type="button"><Sparkles size={15} /><span><strong>Ask LUMA</strong><small>Get personalized insights</small></span><ChevronRight size={15} /></button>
          </div>
          <div className="preferences-list compact-preferences">
            {preferenceItems.slice(0, 3).map((item) => (
              <span key={item}>{item}</span>
            ))}
          </div>
        </GlassPanel>
      </div>

      {editOpen && (
        <div className="profile-modal-overlay">
          <motion.div className="glass-panel profile-edit-modal" initial={{ opacity: 0, y: 16 }} animate={{ opacity: 1, y: 0 }}>
            <div className="modal-header-row">
              <div>
                <h3>Edit Profile</h3>
                <p>Update identity, study focus, current goal, and privacy shortcuts.</p>
              </div>
              <button aria-label="Close" onClick={() => setEditOpen(false)} type="button"><X size={16} /></button>
            </div>

            <div className="profile-edit-grid">
              <section className="profile-edit-block">
                <h4>Identity</h4>
                <div className="profile-edit-avatar-row">
                  <span className={`profile-photo ${draftProfile.avatarStyle === "photo" ? "has-image" : ""}`} style={draftAvatarStyle}>
                    {draftProfile.avatarStyle === "photo" ? "" : draftAvatarText}
                  </span>
                  <div className="profile-avatar-controls">
                    <div className="profile-avatar-mode-row">
                      {(["gradient", "monogram", "photo"] as const).map((mode) => (
                        <button
                          key={mode}
                          className={draftProfile.avatarStyle === mode ? "active" : ""}
                          onClick={() => setDraftProfile((current) => ({ ...current, avatarStyle: mode }))}
                          type="button"
                        >
                          {mode}
                        </button>
                      ))}
                    </div>
                    <div className="profile-gradient-picker">
                      {profileAvatarGradients.map((gradient) => (
                        <button
                          key={gradient}
                          className={draftProfile.avatarGradient === gradient ? "active" : ""}
                          onClick={() => setDraftProfile((current) => ({ ...current, avatarStyle: "gradient", avatarGradient: gradient }))}
                          style={{ background: gradient }}
                          type="button"
                        />
                      ))}
                    </div>
                    <label className="profile-upload-button">
                      <Upload size={14} />
                      Upload Image
                      <input
                        accept="image/*"
                        hidden
                        onChange={(event) => {
                          const file = event.currentTarget.files?.[0];
                          if (!file) return;
                          const reader = new FileReader();
                          reader.onload = () => {
                            setDraftProfile((current) => ({
                              ...current,
                              avatarStyle: "photo",
                              avatarUrl: typeof reader.result === "string" ? reader.result : current.avatarUrl,
                              avatarCrop: current.avatarCrop ?? { zoom: 100, x: 50, y: 50 },
                            }));
                          };
                          reader.readAsDataURL(file);
                        }}
                        type="file"
                      />
                    </label>
                    <button
                      onClick={() => setDraftProfile((current) => ({ ...current, avatarStyle: "monogram", avatarUrl: undefined, avatarCrop: undefined }))}
                      type="button"
                    >
                      Remove Photo
                    </button>
                    {draftProfile.avatarStyle === "photo" && (
                      <div className="profile-crop-controls">
                        {!draftProfile.avatarUrl && <small>Upload a photo to crop.</small>}
                        <label>
                          Crop Zoom
                          <input
                            type="range"
                            min="100"
                            max="180"
                            disabled={!draftProfile.avatarUrl}
                            value={draftProfile.avatarCrop?.zoom ?? 100}
                            onChange={(event) => setDraftProfile((current) => ({
                              ...current,
                              avatarCrop: { zoom: Number(event.currentTarget.value), x: current.avatarCrop?.x ?? 50, y: current.avatarCrop?.y ?? 50 },
                            }))}
                          />
                        </label>
                        <label>
                          Horizontal
                          <input
                            type="range"
                            min="0"
                            max="100"
                            disabled={!draftProfile.avatarUrl}
                            value={draftProfile.avatarCrop?.x ?? 50}
                            onChange={(event) => setDraftProfile((current) => ({
                              ...current,
                              avatarCrop: { zoom: current.avatarCrop?.zoom ?? 100, x: Number(event.currentTarget.value), y: current.avatarCrop?.y ?? 50 },
                            }))}
                          />
                        </label>
                        <label>
                          Vertical
                          <input
                            type="range"
                            min="0"
                            max="100"
                            disabled={!draftProfile.avatarUrl}
                            value={draftProfile.avatarCrop?.y ?? 50}
                            onChange={(event) => setDraftProfile((current) => ({
                              ...current,
                              avatarCrop: { zoom: current.avatarCrop?.zoom ?? 100, x: current.avatarCrop?.x ?? 50, y: Number(event.currentTarget.value) },
                            }))}
                          />
                        </label>
                      </div>
                    )}
                  </div>
                </div>
                <label>
                  Display Name
                  <input value={draftProfile.name} onChange={(event) => setDraftProfile((current) => ({ ...current, name: event.currentTarget.value }))} />
                </label>
                <label>
                  Username
                  <input value={draftProfile.username ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, username: event.currentTarget.value }))} />
                </label>
                <label>
                  School
                  <input value={draftProfile.school ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, school: event.currentTarget.value }))} />
                </label>
                <label>
                  Grade / Year
                  <input value={draftProfile.gradeOrYear ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, gradeOrYear: event.currentTarget.value }))} />
                </label>
                <label>
                  Program / Major
                  <input value={draftProfile.program ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, program: event.currentTarget.value }))} />
                </label>
                <label>
                  Personal Statement
                  <textarea value={draftProfile.bio ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, bio: event.currentTarget.value }))} />
                </label>
                <label>
                  Location
                  <input value={draftProfile.location ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, location: event.currentTarget.value }))} />
                </label>
                <label>
                  Timezone
                  <input value={draftProfile.timezone ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, timezone: event.currentTarget.value }))} />
                </label>
                <label>
                  Language
                  <input value={draftProfile.language ?? ""} onChange={(event) => setDraftProfile((current) => ({ ...current, language: event.currentTarget.value }))} />
                </label>
              </section>

              <section className="profile-edit-block">
                <h4>Study Focus</h4>
                <div className="profile-focus-editor">
                  {draftProfile.focusSubjects.map((subject, index) => (
                    <div key={subject} className="profile-focus-editor-row">
                      <strong>{subject}</strong>
                      <div>
                        <button disabled={index === 0} onClick={() => setDraftProfile((current) => ({ ...current, focusSubjects: moveItem(current.focusSubjects, index, index - 1) }))} type="button">Up</button>
                        <button disabled={index === draftProfile.focusSubjects.length - 1} onClick={() => setDraftProfile((current) => ({ ...current, focusSubjects: moveItem(current.focusSubjects, index, index + 1) }))} type="button">Down</button>
                        <button onClick={() => setDraftProfile((current) => ({ ...current, focusSubjects: current.focusSubjects.filter((item) => item !== subject) }))} type="button">Remove</button>
                      </div>
                    </div>
                  ))}
                </div>
                <div className="profile-add-focus">
                  <input
                    placeholder="Add focus area"
                    value={focusInput}
                    onChange={(event) => setFocusInput(event.currentTarget.value)}
                  />
                  <button onClick={addFocusArea} type="button">Add</button>
                </div>
                <small>Pin up to 6 focus areas from subjects, projects, or manual entries.</small>

                <h4>Visible Metrics</h4>
                <div className="profile-metric-picker">
                  {profileMetricDefinitions.map((metric) => (
                    <button
                      key={metric.key}
                      className={draftProfile.metricsVisible.includes(metric.key) ? "active" : ""}
                      onClick={() => toggleMetric(metric.key)}
                      type="button"
                    >
                      {metric.label}
                    </button>
                  ))}
                </div>
                <small>Choose up to 4 metrics for the default Profile view.</small>
              </section>

              <section className="profile-edit-block">
                <h4>Current Goal</h4>
                <div className="profile-goal-type-grid">
                  {profileGoalTypeOptions.map((option) => (
                    <button
                      key={option.value}
                      className={draftProfile.currentGoal?.type === option.value ? "active" : ""}
                      onClick={() => updateDraftGoal({ type: option.value })}
                      type="button"
                    >
                      {option.label}
                    </button>
                  ))}
                </div>
                <label>
                  Goal Title
                  <input value={draftProfile.currentGoal?.title ?? ""} onChange={(event) => updateDraftGoal({ title: event.currentTarget.value })} />
                </label>
                <label>
                  Deadline
                  <input type="datetime-local" value={draftProfile.currentGoal?.dueAt ? format(parseISO(draftProfile.currentGoal.dueAt), "yyyy-MM-dd'T'HH:mm") : ""} onChange={(event) => updateDraftGoal({ dueAt: event.currentTarget.value })} />
                </label>
                <label>
                  Target Focus Minutes
                  <input type="number" min="0" value={draftProfile.currentGoal?.targetFocusMinutes ?? 0} onChange={(event) => updateDraftGoal({ targetFocusMinutes: Number(event.currentTarget.value) })} />
                </label>
                <label>
                  Linked Subject
                  <select value={draftProfile.currentGoal?.linkedSubjectId ?? ""} onChange={(event) => updateDraftGoal({ linkedSubjectId: event.currentTarget.value })}>
                    {snapshot.subjects.map((subject) => (
                      <option key={subject.id} value={subject.id}>{subject.name}</option>
                    ))}
                  </select>
                </label>
                <label>
                  Progress
                  <input type="range" min="0" max="100" value={draftProfile.currentGoal?.progress ?? 0} onChange={(event) => updateDraftGoal({ progress: Number(event.currentTarget.value) })} />
                  <strong>{draftProfile.currentGoal?.progress ?? 0}% complete</strong>
                </label>
                <div className="goal-ai-plan">
                  <strong>Your goal can be completed in {Math.max(1, Math.ceil((draftProfile.currentGoal?.targetFocusMinutes ?? 300) / Math.max(snapshot.settings.focusSessionMinutes, 1)))} sessions.</strong>
                  <small>Suggested flow: weekly study blocks, one review pass, and a final checkpoint before the deadline.</small>
                </div>
              </section>

              <section className="profile-edit-block">
                <h4>Privacy Controls</h4>
                {([
                  ["profile", "Profile"],
                  ["avatar", "Avatar"],
                  ["studyFocus", "Study Focus"],
                ] as const).map(([key, label]) => (
                  <div key={key} className="profile-privacy-control">
                    <strong>{label}</strong>
                    <div className="profile-segmented">
                      {(["private", "friends", "groups"] as const).map((value) => (
                        <button
                          key={value}
                          className={draftProfile.visibility[key] === value ? "active" : ""}
                          onClick={() => setDraftProfile((current) => ({ ...current, visibility: { ...current.visibility, [key]: value } }))}
                          type="button"
                        >
                          {value}
                        </button>
                      ))}
                    </div>
                  </div>
                ))}
                <div className="profile-privacy-note">
                  <Lock size={15} />
                  <span>Public profile stays off by default. Most profile data remains private unless you choose otherwise.</span>
                </div>
              </section>
            </div>

            <div className="modal-action-buttons">
              <button className="modal-cancel-btn" onClick={() => setEditOpen(false)} type="button">Cancel</button>
              <button className="modal-submit-btn" onClick={() => void saveProfileDraft()} type="button">Save Profile</button>
            </div>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function UniversalHeader({ title, subtitle, avatar }: { title: string; subtitle: string; avatar: string }) {
  return (
    <header className="universal-header">
      <div>
        <h1>{title} <Sparkles size={25} /></h1>
        <p>{subtitle}</p>
      </div>
      <div className="home-user-actions">
        <button type="button" aria-label="Notifications"><Bell size={22} /></button>
        <span className="home-avatar" aria-label="Profile">{avatar.slice(0, 2)}<i aria-hidden="true" /></span>
      </div>
    </header>
  );
}

function SettingsScreen({
  snapshot,
  saveSettings,
  reset,
  signOut,
}: {
  snapshot: LumaSnapshot;
  saveSettings: (settings: UserSettings) => Promise<void>;
  reset: () => Promise<void>;
  signOut: () => Promise<void>;
}) {
  const [aiStatus, setAiStatus] = useState("checking");
  const [activeCategory, setActiveCategory] = useState<SettingsCategory>("Appearance");
  const [browserSupport, setBrowserSupport] = useState<BrowserLocalSupport | null>(null);
  const [browserModelCached, setBrowserModelCached] = useState(false);
  const [browserModelProgress, setBrowserModelProgress] = useState("");
  const [browserModelLoading, setBrowserModelLoading] = useState(false);

  useEffect(() => {
    if (snapshot.settings.localAiEndpoint === "browser-webllm") {
      void Promise.resolve().then(() => {
        const runtime = getBrowserLocalRuntimeState();
        setAiStatus(runtime.ready ? "ready" : "not loaded");
      });
      void checkBrowserLocalSupport().then(setBrowserSupport);
      void checkBrowserModelCached(snapshot.settings.localAiModel).then(setBrowserModelCached);
      return;
    }
    const provider = new OllamaProvider(snapshot.settings.localAiEndpoint, snapshot.settings.localAiModel);
    void provider.status().then(setAiStatus);
  }, [snapshot.settings.localAiEndpoint, snapshot.settings.localAiModel]);

  const loadBrowserModel = async () => {
    setBrowserModelLoading(true);
    setBrowserModelProgress("Preparing browser local model...");
    try {
      await initializeBrowserLocalModel(snapshot.settings.localAiModel, (progress) => {
        setBrowserModelProgress(`${Math.round(progress.progress * 100)}% · ${progress.message}`);
      });
      setAiStatus("ready");
      setBrowserModelCached(await checkBrowserModelCached(snapshot.settings.localAiModel));
      setBrowserModelProgress("Local model ready. LUMA Agent will use it for answers.");
    } catch (error) {
      setAiStatus("error");
      setBrowserModelProgress(error instanceof Error ? error.message : String(error));
    } finally {
      setBrowserModelLoading(false);
    }
  };

  const unloadBrowserModel = async () => {
    await unloadBrowserLocalModel();
    setAiStatus("not loaded");
    setBrowserModelProgress("Model unloaded. Cached weights stay on this device.");
  };

  const purgeBrowserModel = async () => {
    if (!confirm("Delete cached model weights from this browser? LUMA can download them again later.")) return;
    setBrowserModelLoading(true);
    try {
      await deleteBrowserModelCache(snapshot.settings.localAiModel);
      setBrowserModelCached(false);
      setAiStatus("not loaded");
      setBrowserModelProgress("Cached model weights deleted.");
    } finally {
      setBrowserModelLoading(false);
    }
  };

  const updateSettings = (settings: UserSettings) => {
    void saveSettings(settings);
  };

  const patchSettings = (patch: Partial<UserSettings>) => {
    updateSettings({ ...snapshot.settings, ...patch });
  };

  const exportMyData = () => {
    downloadBlob(
      new Blob([JSON.stringify(snapshot, null, 2)], { type: "application/json;charset=utf-8" }),
      `luma-export-${new Date().toISOString().slice(0, 10)}.json`,
    );
  };

  const openPrivacyPolicy = () => {
    const summary = [
      "LUMA Privacy Summary",
      "",
      "1. Study data is stored locally in your browser by default.",
      "2. Cloud sync stays off until you enable it.",
      "3. AI memory only stores preferences and study patterns when enabled.",
      "4. Diagnostics are anonymous and can be disabled at any time.",
      "5. Export lets you download your current local snapshot as JSON.",
    ].join("\n");
    window.open(`data:text/plain;charset=utf-8,${encodeURIComponent(summary)}`, "_blank", "noopener,noreferrer");
  };

  const resetPreferences = () => {
    if (!confirm("Reset settings and home customization to defaults? Your study data will stay untouched.")) return;
    updateSettings(createSeedData().settings);
  };

  const renderToggle = (
    label: string,
    description: string,
    checked: boolean,
    onChange: (checked: boolean) => void,
    Icon: LucideIcon = Shield,
  ) => (
    <label className="luma-toggle-row">
      <span className="toggle-copy">
        <Icon size={17} />
        <span>
          <strong>{label}</strong>
          <small>{description}</small>
        </span>
      </span>
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.currentTarget.checked)} />
    </label>
  );

  const renderAppearancePanel = () => (
    <GlassPanel className="appearance-panel settings-primary-panel">
      <div className="settings-panel-heading">
        <h3>Appearance / Home Customization</h3>
        <p>Personalize LUMA to match your space and flow.</p>
      </div>
      <h4>Theme</h4>
      <div className="scene-choice-row">
        {backgroundScenes.map((scene) => (
          <button
            key={scene.id}
            className={snapshot.settings.backgroundScene === scene.id ? "selected" : ""}
            onClick={() => patchSettings({ backgroundScene: scene.id, theme: scene.id === "Night City" ? "Night Bloom" : scene.id })}
            style={{ "--scene-thumb": `url("${scene.file}")` } as CSSProperties}
            type="button"
          >
            <span>{snapshot.settings.backgroundScene === scene.id && <Check size={14} />}</span>
            <strong>{scene.id === "Night City" ? "Night Bloom" : scene.id}</strong>
          </button>
        ))}
      </div>
      <h4>Accent Color</h4>
      <div className="accent-row">
        {accentChoices.map((color) => (
          <button
            key={color}
            aria-label={color}
            className={snapshot.settings.accentColor === color ? "selected" : ""}
            onClick={() => patchSettings({ accentColor: color })}
            style={{ background: color }}
            type="button"
          />
        ))}
      </div>
      <label className="settings-slider">
        <span>Blur Strength<small>Adjust the background blur across LUMA.</small></span>
        <button type="button" onClick={() => patchSettings({ blur: Math.max(0, snapshot.settings.blur - 8) })}><Minus size={14} /></button>
        <input min="0" max="100" type="range" value={snapshot.settings.blur} onChange={(event) => patchSettings({ blur: Number(event.currentTarget.value) })} />
        <button type="button" onClick={() => patchSettings({ blur: Math.min(100, snapshot.settings.blur + 8) })}><Plus size={14} /></button>
        <strong>{snapshot.settings.blur}%</strong>
      </label>
      <h4>Widget Style</h4>
      <div className="widget-style-row">
        {widgetStyles.map((style) => (
          <button key={style.value} className={snapshot.settings.widgetStyle === style.value ? "selected" : ""} onClick={() => patchSettings({ widgetStyle: style.value })} type="button">
            <span>{snapshot.settings.widgetStyle === style.value && <Check size={13} />}</span>
            <strong>{style.label}</strong>
            <small>{style.description}</small>
          </button>
        ))}
      </div>
      <h4>Layout Density</h4>
      <div className="settings-segmented">
        {(["Compact", "Comfortable", "Spacious"] as const).map((item) => (
          <button key={item} className={snapshot.settings.density === item ? "active" : ""} onClick={() => patchSettings({ density: item })} type="button">
            {item}
          </button>
        ))}
      </div>
    </GlassPanel>
  );

  const renderPrimaryPanel = () => {
    if (activeCategory === "Appearance") {
      return renderAppearancePanel();
    }

    if (activeCategory === "General") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>General Settings</h3>
            <p>Language, time, launch behavior, and core app preferences.</p>
          </div>
          <div className="settings-form-grid">
            <label><Languages size={17} /> Language<select value={snapshot.settings.language} onChange={(event) => patchSettings({ language: event.currentTarget.value as UserSettings["language"] })}><option>English</option><option>Thai</option><option>Japanese</option><option>System Default</option></select></label>
            <label><Clock size={17} /> Time Format<select value={snapshot.settings.timeFormat} onChange={(event) => patchSettings({ timeFormat: event.currentTarget.value as UserSettings["timeFormat"] })}><option>12-hour</option><option>24-hour</option></select></label>
            <label><CalendarDays size={17} /> Start Day<select value={snapshot.settings.weekStart} onChange={(event) => patchSettings({ weekStart: event.currentTarget.value as UserSettings["weekStart"] })}><option>Monday</option><option>Sunday</option><option>Saturday</option></select></label>
            <label><Home size={17} /> Default Landing<select value={snapshot.settings.landingPage} onChange={(event) => patchSettings({ landingPage: event.currentTarget.value as UserSettings["landingPage"] })}><option>Home</option><option>Calendar</option><option>Learn</option><option>Last Opened Page</option></select></label>
          </div>
          <div className="settings-stack">
            {renderToggle("Auto-start LUMA", "Launch LUMA when your device starts.", snapshot.settings.autoStart, (checked) => patchSettings({ autoStart: checked }), Zap)}
            {renderToggle("Minimize to dock", "Keep LUMA running quietly in the background.", snapshot.settings.minimizeToDock, (checked) => patchSettings({ minimizeToDock: checked }), PanelTop)}
            {renderToggle("Focus mode on launch", "Open the app in a focus-first state.", snapshot.settings.focusOnLaunch, (checked) => patchSettings({ focusOnLaunch: checked }), Play)}
          </div>
          <button className="danger-neutral-button" type="button" onClick={resetPreferences}>
            <RotateCcw size={16} /> Reset Preferences
          </button>
        </GlassPanel>
      );
    }

    if (activeCategory === "Focus") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Focus Settings</h3>
            <p>Control deep work timing, motion, and sound behavior.</p>
          </div>
          <div className="settings-form-grid">
            <label><Timer size={17} /> Focus Session<select value={String(snapshot.settings.focusSessionMinutes)} onChange={(event) => patchSettings({ focusSessionMinutes: Number(event.currentTarget.value) })}><option value="25">25 min</option><option value="45">45 min</option><option value="60">60 min</option><option value="90">90 min</option></select></label>
            <label><Volume2 size={17} /> Focus Sound<select value={snapshot.settings.focusSound} onChange={(event) => patchSettings({ focusSound: event.currentTarget.value as UserSettings["focusSound"] })}><option>Rain</option><option>Ocean</option><option>Forest</option><option>None</option></select></label>
            <label><Clock size={17} /> Short Break<select value={String(snapshot.settings.shortBreakMinutes)} onChange={(event) => patchSettings({ shortBreakMinutes: Number(event.currentTarget.value) })}><option value="5">5 min</option><option value="10">10 min</option><option value="15">15 min</option></select></label>
            <label><Sparkles size={17} /> Motion<select value={snapshot.settings.motion} onChange={(event) => patchSettings({ motion: event.currentTarget.value as UserSettings["motion"] })}><option value="low">Low</option><option value="balanced">Balanced</option><option value="expressive">Expressive</option></select></label>
          </div>
          <div className="settings-stack">
            {renderToggle("Auto-start breaks", "Start the short break timer automatically after a focus session.", snapshot.settings.focusAutoStartBreaks, (checked) => patchSettings({ focusAutoStartBreaks: checked }), Timer)}
            {renderToggle("Auto-start next session", "Roll into the next focus block without extra taps.", snapshot.settings.focusAutoStartSessions, (checked) => patchSettings({ focusAutoStartSessions: checked }), RefreshCw)}
          </div>
        </GlassPanel>
      );
    }

    if (activeCategory === "Notifications") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Notifications</h3>
            <p>Keep reminders useful and quiet when you need uninterrupted study time.</p>
          </div>
          <div className="settings-stack">
            {renderToggle("Notifications enabled", "Allow LUMA to surface reminders and prompts.", snapshot.settings.notificationsEnabled, (checked) => patchSettings({ notificationsEnabled: checked }), Bell)}
            {renderToggle("Deadline reminders", "Get nudges before homework, projects, and exams.", snapshot.settings.deadlineReminders, (checked) => patchSettings({ deadlineReminders: checked }), CalendarDays)}
            {renderToggle("Daily agenda", "Show a daily summary of today’s priorities.", snapshot.settings.dailyAgenda, (checked) => patchSettings({ dailyAgenda: checked }), BookOpen)}
            {renderToggle("Sound effects", "Play subtle interface sounds for confirmations and alerts.", snapshot.settings.soundEffects, (checked) => patchSettings({ soundEffects: checked }), Headphones)}
            {renderToggle("Quiet hours", "Pause non-urgent reminders during your quiet window.", snapshot.settings.quietHoursEnabled, (checked) => patchSettings({ quietHoursEnabled: checked }), Bell)}
          </div>
          <div className="settings-form-grid">
            <label><Clock size={17} /> Quiet Starts<input type="time" value={snapshot.settings.quietHoursStart} onChange={(event) => patchSettings({ quietHoursStart: event.currentTarget.value })} /></label>
            <label><Clock size={17} /> Quiet Ends<input type="time" value={snapshot.settings.quietHoursEnd} onChange={(event) => patchSettings({ quietHoursEnd: event.currentTarget.value })} /></label>
          </div>
        </GlassPanel>
      );
    }

    if (activeCategory === "Widgets") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Widgets</h3>
            <p>Control how the home command space behaves and how widgets are arranged.</p>
          </div>
          <div className="settings-form-grid">
            <label><LayoutGrid size={17} /> Experience Preset<select value={snapshot.settings.experiencePreset} onChange={(event) => patchSettings(applyStarterPreset(snapshot.settings, event.currentTarget.value as ExperiencePreset))}><option>Balanced</option><option>Exam Sprint</option><option>Project Studio</option><option>Minimal Focus</option></select></label>
            <label><Palette size={17} /> Widget Style<select value={snapshot.settings.widgetStyle} onChange={(event) => patchSettings({ widgetStyle: event.currentTarget.value as UserSettings["widgetStyle"] })}><option>Glassmorphism</option><option>Soft Glow</option><option>Minimal</option></select></label>
          </div>
          <div className="settings-stack">
            {renderToggle("Snap widgets to grid", "Keep home widgets aligned neatly while rearranging.", snapshot.settings.widgetsSnapToGrid, (checked) => patchSettings({ widgetsSnapToGrid: checked }), LayoutGrid)}
            {renderToggle("Show widget labels", "Keep titles visible inside the home layout.", snapshot.settings.showWidgetLabels, (checked) => patchSettings({ showWidgetLabels: checked }), Info)}
          </div>
        </GlassPanel>
      );
    }

    if (activeCategory === "AI Privacy") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Your study space stays yours.</h3>
            <p>LUMA processes study data locally by default. You choose what can be remembered, synced, or shared.</p>
          </div>
          <div className="settings-stack">
            {renderToggle("Local-only data processing", "Keep study material and AI processing on this device.", snapshot.settings.localOnlyMaterials, (checked) => patchSettings({ localOnlyMaterials: checked }), Lock)}
            {renderToggle("AI Memory & Context", "Allow LUMA to remember preferences, schedule patterns, and study habits.", snapshot.settings.aiMemoryEnabled, (checked) => patchSettings({ aiMemoryEnabled: checked }), Bot)}
            {renderToggle("Group AI Access", "Use shared group files and tasks as context inside group workspaces.", snapshot.settings.groupAiAccess, (checked) => patchSettings({ groupAiAccess: checked }), Users)}
            {renderToggle("Cloud Sync", "Sync encrypted LUMA data across your signed-in devices.", snapshot.settings.allowCloudSync, (checked) => patchSettings({ allowCloudSync: checked }), Cloud)}
            {renderToggle("Diagnostics", "Share anonymous performance data to improve LUMA.", snapshot.settings.diagnostics, (checked) => patchSettings({ diagnostics: checked }), Gauge)}
          </div>
          <div className="button-row">
            <ActionButton tone="ghost" onClick={exportMyData}><Download size={15} /> Export My Data</ActionButton>
            <ActionButton tone="danger" onClick={() => confirm("Delete local AI memory and model context?") && patchSettings({ aiMemoryEnabled: false, groupAiAccess: false })}><Trash2 size={15} /> Delete Local AI Data</ActionButton>
          </div>
        </GlassPanel>
      );
    }

    if (activeCategory === "Local Model") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Local Model Settings</h3>
            <p>Choose how LUMA runs private study intelligence on this device.</p>
          </div>
          <div className="model-status-card">
            <Orb size="medium" />
            <div>
              <strong>LUMA Local Model</strong>
              <span>{aiStatus}</span>
              <small>{snapshot.settings.localAiModel} · {snapshot.settings.localAiPerformance}</small>
            </div>
          </div>
          <div className="settings-form-grid">
            <label><Bot size={17} /> Runtime<select value={snapshot.settings.localAiEndpoint} onChange={(event) => patchSettings({ localAiEndpoint: event.currentTarget.value, localAiModel: event.currentTarget.value === "browser-webllm" ? BROWSER_LOCAL_MODELS[0].id : "llama3.2:latest" })}><option value="browser-webllm">Browser WebLLM</option><option value="http://localhost:11434">Ollama localhost</option></select></label>
            {snapshot.settings.localAiEndpoint === "browser-webllm" ? (
              <label><Bot size={17} /> Model<select value={snapshot.settings.localAiModel} onChange={(event) => patchSettings({ localAiModel: event.currentTarget.value })}>{BROWSER_LOCAL_MODELS.map((model) => <option key={model.id} value={model.id}>{model.name}</option>)}</select></label>
            ) : (
              <label><Bot size={17} /> Model<select value={snapshot.settings.localAiModel} onChange={(event) => patchSettings({ localAiModel: event.currentTarget.value })}><option value="llama3.2:latest">Llama 3.2 Latest</option><option value="qwen2.5:7b">Qwen 2.5 7B</option><option value="mistral:7b">Mistral 7B</option><option value="phi3:mini">Phi 3 Mini</option></select></label>
            )}
            <label><Gauge size={17} /> Performance<select value={snapshot.settings.localAiPerformance} onChange={(event) => patchSettings({ localAiPerformance: event.currentTarget.value as UserSettings["localAiPerformance"] })}><option>Balanced</option><option>Battery Saver</option><option>Fastest</option><option>Best Quality</option></select></label>
            {snapshot.settings.localAiEndpoint !== "browser-webllm" && <label><Globe2 size={17} /> Endpoint<input value={snapshot.settings.localAiEndpoint} onChange={(event) => patchSettings({ localAiEndpoint: event.currentTarget.value })} placeholder="http://localhost:11434" /></label>}
          </div>
          {snapshot.settings.localAiEndpoint === "browser-webllm" && (
            <div className="local-model-console">
              <div className="storage-meter"><span><HardDrive size={16} /> Browser Model Cache</span><strong>{browserModelCached ? "Downloaded on this device" : "Not downloaded yet"}</strong><i /></div>
              <p>{browserSupport?.supported ? "WebGPU is ready. Loading will download weights once, then cache them locally." : browserSupport?.reason ?? "Checking WebGPU support..."}</p>
              {browserModelProgress && <small>{browserModelProgress}</small>}
              <div className="button-row">
                <ActionButton tone="primary" onClick={() => void loadBrowserModel()} disabled={browserModelLoading || browserSupport?.supported === false}><Download size={15} /> {browserModelCached ? "Load Local Model" : "Download & Load"}</ActionButton>
                <ActionButton tone="ghost" onClick={() => void unloadBrowserModel()} disabled={browserModelLoading || !getBrowserLocalRuntimeState().ready}><RotateCcw size={15} /> Unload</ActionButton>
                <ActionButton tone="danger" onClick={() => void purgeBrowserModel()} disabled={browserModelLoading || !browserModelCached}><Trash2 size={15} /> Delete Weights</ActionButton>
              </div>
            </div>
          )}
          {snapshot.settings.localAiEndpoint !== "browser-webllm" && <div className="storage-meter"><span><HardDrive size={16} /> AI Storage</span><strong>External Ollama runtime</strong><i /></div>}
        </GlassPanel>
      );
    }

    if (activeCategory === "Backup & Sync") {
      return (
        <GlassPanel className="appearance-panel settings-primary-panel">
          <div className="settings-panel-heading">
            <h3>Backup & Sync</h3>
            <p>Control sync, exports, and full local restore behavior.</p>
          </div>
          <div className="settings-form-grid">
            <label><RefreshCw size={17} /> Backup Frequency<select value={snapshot.settings.backupFrequency} onChange={(event) => patchSettings({ backupFrequency: event.currentTarget.value as UserSettings["backupFrequency"] })}><option>Daily</option><option>Weekly</option><option>Manual</option></select></label>
            <label><Cloud size={17} /> Sync Mode<select value={snapshot.settings.allowCloudSync ? "Enabled" : "Disabled"} onChange={(event) => patchSettings({ allowCloudSync: event.currentTarget.value === "Enabled" })}><option>Disabled</option><option>Enabled</option></select></label>
          </div>
          <div className="button-row">
            <ActionButton tone="ghost" onClick={exportMyData}><Download size={15} /> Download Backup</ActionButton>
            <ActionButton tone="ghost" onClick={() => confirm("Sign out of this local LUMA session? Your data stays on this device.") && void signOut()}><Lock size={15} /> Sign Out</ActionButton>
            <ActionButton tone="danger" onClick={() => confirm("Reset to the M.5 IM timetable preset? This clears local tasks, materials, groups, and settings.") && void reset()}><RotateCcw size={15} /> Reset to M.5 IM Timetable</ActionButton>
          </div>
        </GlassPanel>
      );
    }

    return (
      <GlassPanel className="appearance-panel settings-primary-panel">
        <div className="settings-panel-heading">
          <h3>About LUMA</h3>
          <p>Version, data model, and privacy shortcuts.</p>
        </div>
        <div className="settings-placeholder-grid">
          <button type="button"><Info size={16} /><strong>Version 0.1.0</strong><small>Local-first student workspace</small></button>
          <button type="button"><Shield size={16} /><strong>Privacy First</strong><small>AI and study data stay local by default</small></button>
          <button type="button"><HardDrive size={16} /><strong>IndexedDB Storage</strong><small>Runs directly in your browser</small></button>
          <button type="button"><Sparkles size={16} /><strong>{snapshot.settings.localAiModel}</strong><small>Current local model</small></button>
        </div>
        <div className="button-row">
          <ActionButton tone="ghost" onClick={openPrivacyPolicy}>View Privacy Policy <Share2 size={13} /></ActionButton>
          <ActionButton onClick={exportMyData}><Download size={15} /> Export Snapshot</ActionButton>
        </div>
      </GlassPanel>
    );
  };

  return (
    <div className="universal-screen settings-universal">
      <UniversalHeader title="Settings" subtitle="Designed around you." avatar={snapshot.profile.avatar} />
      <div className="settings-dashboard">
        <GlassPanel className="settings-menu">
          {settingsMenuItems.map((item) => {
            const Icon = item.icon;
            return (
              <button key={item.title} className={activeCategory === item.title ? "active" : ""} onClick={() => setActiveCategory(item.title)} type="button">
                <Icon size={20} />
                <strong>{item.title}</strong>
                <small>{item.description}</small>
                {activeCategory === item.title && <Sparkles className="settings-active-spark" size={14} />}
              </button>
            );
          })}
        </GlassPanel>
        {renderPrimaryPanel()}
        <GlassPanel className="privacy-panel settings-privacy">
          <h3>Quick Status</h3>
          <p>Your current setup updates live as you change settings.</p>
          <div className="local-status">
            <Orb size="small" />
            <div><strong>LUMA Local Model</strong><span className="active-status">{aiStatus}</span><small>{snapshot.settings.localAiModel} · {snapshot.settings.localAiPerformance}</small></div>
            <ActionButton tone="ghost" onClick={() => setActiveCategory("Local Model")}>Manage Model</ActionButton>
          </div>
          <div className="settings-stack compact">
            {renderToggle("Local-only data processing", "Keep all data on this device.", snapshot.settings.localOnlyMaterials, (checked) => patchSettings({ localOnlyMaterials: checked }), Shield)}
            {renderToggle("Memory & context", "Allow LUMA to remember preferences.", snapshot.settings.aiMemoryEnabled, (checked) => patchSettings({ aiMemoryEnabled: checked }), Lock)}
            {renderToggle("Cross-device sync", "Sync settings and local memory.", snapshot.settings.allowCloudSync, (checked) => patchSettings({ allowCloudSync: checked }), RefreshCw)}
          </div>
          <button className="privacy-link-button" type="button" onClick={openPrivacyPolicy}>View Privacy Policy <Share2 size={13} /></button>
        </GlassPanel>
        <GlassPanel className="general-settings-panel">
          <h3>Current Preferences</h3>
          <div className="settings-grid">
            <label><Globe2 size={16} /> Language<select value={snapshot.settings.language} onChange={(event) => patchSettings({ language: event.currentTarget.value as UserSettings["language"] })}><option>English</option><option>Thai</option><option>Japanese</option><option>System Default</option></select></label>
            <label><Clock size={16} /> Time format<select value={snapshot.settings.timeFormat} onChange={(event) => patchSettings({ timeFormat: event.currentTarget.value as UserSettings["timeFormat"] })}><option>24-hour</option><option>12-hour</option></select></label>
          </div>
          <div className="settings-stack compact">
            {renderToggle("Auto-start LUMA", "Launch LUMA when system starts.", snapshot.settings.autoStart, (checked) => patchSettings({ autoStart: checked }), Smartphone)}
            {renderToggle("Minimize to dock", "Keep LUMA running in the background.", snapshot.settings.minimizeToDock, (checked) => patchSettings({ minimizeToDock: checked }), PanelTop)}
            {renderToggle("Sound effects", "Play soft UI sounds.", snapshot.settings.soundEffects, (checked) => patchSettings({ soundEffects: checked }), Headphones)}
            {renderToggle("Share diagnostics", "Help improve LUMA privately.", snapshot.settings.diagnostics, (checked) => patchSettings({ diagnostics: checked }), Gauge)}
          </div>
        </GlassPanel>
      </div>
    </div>
  );
}

function AgentMessageContent({ content }: { content: string }) {
  const blocks: Array<{ type: "paragraph"; text: string } | { type: "ul" | "ol"; items: string[] }> = [];
  let list: { type: "ul" | "ol"; items: string[] } | null = null;

  const flushList = () => {
    if (list) {
      blocks.push(list);
      list = null;
    }
  };

  content
    .split("\n")
    .map((line) => line.trim())
    .forEach((line) => {
      if (!line) {
        flushList();
        return;
      }
      const ordered = line.match(/^\d+[.)]\s+(.+)$/);
      const unordered = line.match(/^[-*]\s+(.+)$/);
      if (ordered || unordered) {
        const type = ordered ? "ol" : "ul";
        const item = ordered?.[1] ?? unordered?.[1] ?? "";
        if (!list || list.type !== type) {
          flushList();
          list = { type, items: [] };
        }
        list.items.push(item);
        return;
      }
      flushList();
      blocks.push({ type: "paragraph", text: line });
    });
  flushList();

  return (
    <div className="agent-message-body">
      {blocks.map((block, index) =>
        block.type === "paragraph" ? (
          <p key={`${block.type}-${index}`}>{block.text}</p>
        ) : block.type === "ol" ? (
          <ol key={`${block.type}-${index}`}>
            {block.items.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ol>
        ) : (
          <ul key={`${block.type}-${index}`}>
            {block.items.map((item) => (
              <li key={item}>{item}</li>
            ))}
          </ul>
        ),
      )}
    </div>
  );
}

function GlobalAgent({
  snapshot,
  agentResult,
  setAgentResult,
  executeAction,
  route,
  navigate,
  refresh,
}: {
  snapshot: LumaSnapshot;
  agentResult: AgentResult | null;
  setAgentResult: (result: AgentResult | null) => void;
  executeAction: (action: AgentAction) => Promise<void>;
  route: RouteId;
  navigate: (route: RouteId, target?: Omit<AgentFocusTarget, "route">) => void;
  refresh: () => Promise<LumaSnapshot>;
}) {
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [pendingAction, setPendingAction] = useState<AgentAction | null>(null);
  const [expanded, setExpanded] = useState(false);
  const [listening, setListening] = useState(false);
  const [pendingPrompt, setPendingPrompt] = useState("");
  const inputRef = useRef<HTMLInputElement | null>(null);
  const transcriptRef = useRef<HTMLDivElement | null>(null);
  const recentPrompts = snapshot.agentMessages
    .filter((message) => message.role === "user")
    .sort((a, b) => parseISO(b.createdAt).getTime() - parseISO(a.createdAt).getTime())
    .slice(0, 3);
  const activeConversation = Boolean(agentResult || expanded || busy);
  const suggestions = [
    { label: "What is due this week?", prompt: "What is due this week?" },
    { label: "Make a study plan", prompt: "Make a study plan around my deadlines and current materials." },
    { label: "Explain binary search", prompt: "Explain binary search from my notes and show the source." },
    { label: "Group progress", prompt: "Summarize our group progress and find unfinished tasks." },
  ];
  const quickContext = [
    route === "home" ? "Home" : route[0].toUpperCase() + route.slice(1),
    `${snapshot.tasks.filter((task) => task.status !== "done").length} active tasks`,
    `${snapshot.materials.length} vault files`,
    snapshot.settings.aiMemoryEnabled ? "Memory on" : "Memory off",
  ];
  const currentScope = route === "together" ? "group" : "personal";
  const latestConversation = snapshot.agentConversations
    .filter((conversation) => conversation.scope === currentScope || currentScope === "personal")
    .sort((a, b) => parseISO(b.updatedAt).getTime() - parseISO(a.updatedAt).getTime())[0];
  const savedThread = latestConversation
    ? snapshot.agentMessages
        .filter((message) => message.conversationId === latestConversation.id)
        .sort((a, b) => parseISO(a.createdAt).getTime() - parseISO(b.createdAt).getTime())
        .slice(-14)
    : [];
  const threadItems = [
    ...savedThread.map((message) => ({
      id: message.id,
      role: message.role,
      content: message.content,
      contextRefs: message.contextRefs,
    })),
    ...(busy && pendingPrompt
      ? [{ id: "pending-user", role: "user" as const, content: pendingPrompt }]
      : []),
  ];
  const latestSavedAssistant = [...savedThread].reverse().find((message) => message.role === "assistant");
  const shouldShowLiveAssistant = agentResult && latestSavedAssistant?.content !== agentResult.answer && !busy;
  const visibleThread = shouldShowLiveAssistant
    ? [
        ...threadItems,
        {
          id: "live-assistant",
          role: "assistant" as const,
          content: agentResult.answer,
          contextRefs: agentResult.contextRefs,
        },
      ]
    : threadItems;

  useEffect(() => {
    const listener = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        setExpanded(true);
        inputRef.current?.focus();
      }
    };
    window.addEventListener("keydown", listener);
    return () => window.removeEventListener("keydown", listener);
  }, []);

  useEffect(() => {
    if (activeConversation) {
      const scrollTranscript = () => {
        const transcript = transcriptRef.current;
        if (transcript) {
          transcript.scrollTop = transcript.scrollHeight;
        }
      };
      window.requestAnimationFrame(scrollTranscript);
      const settleTimer = window.setTimeout(scrollTranscript, 140);
      return () => window.clearTimeout(settleTimer);
    }
    return undefined;
  }, [activeConversation, visibleThread.length, agentResult?.answer, busy]);

  const submit = async () => {
    if (!input.trim()) return;
    const prompt = input.trim();
    setBusy(true);
    setExpanded(true);
    setPendingAction(null);
    setPendingPrompt(prompt);
    try {
      const result = await runLumaAgent(prompt, snapshot, snapshot.settings, route);
      setAgentResult(result);
      await refresh();
      setInput("");
    } catch {
      setAgentResult({
        answer: "I could not complete that request locally. Check your local AI settings or try a simpler action.",
        actions: [],
        provider: "fallback",
        status: "error",
      });
    } finally {
      setBusy(false);
      setPendingPrompt("");
    }
  };

  const chooseAction = (action: AgentAction) => {
    if (action.type !== "open-route" && (action.preview || action.requiresConfirmation)) {
      setPendingAction(action);
      return;
    }
    void executeAction(action);
  };

  const applySuggestedPrompt = (prompt: string) => {
    setInput(prompt);
    setExpanded(true);
    window.setTimeout(() => inputRef.current?.focus(), 0);
  };

  const startVoice = () => {
    const speechWindow = window as typeof window & {
      SpeechRecognition?: new () => {
        lang: string;
        interimResults: boolean;
        start: () => void;
        onresult: ((event: { results: { 0: { 0: { transcript: string } } } }) => void) | null;
        onend: (() => void) | null;
        onerror: (() => void) | null;
      };
      webkitSpeechRecognition?: new () => {
        lang: string;
        interimResults: boolean;
        start: () => void;
        onresult: ((event: { results: { 0: { 0: { transcript: string } } } }) => void) | null;
        onend: (() => void) | null;
        onerror: (() => void) | null;
      };
    };
    const Recognition = speechWindow.SpeechRecognition ?? speechWindow.webkitSpeechRecognition;
    if (!Recognition) {
      setAgentResult({
        answer: "Voice input is not available in this browser. Type the request and I will use the same agent flow.",
        actions: [],
        provider: "fallback",
        status: "offline",
      });
      setExpanded(true);
      return;
    }
    const recognition = new Recognition();
    recognition.lang = "en-US";
    recognition.interimResults = false;
    recognition.onresult = (event) => {
      setInput(event.results[0][0].transcript);
      setExpanded(true);
      inputRef.current?.focus();
    };
    recognition.onend = () => setListening(false);
    recognition.onerror = () => setListening(false);
    setListening(true);
    recognition.start();
  };

  const runUndo = async () => {
    const actionId = agentResult?.undo?.actionId;
    if (!actionId) return;
    const result = await undoAction(actionId);
    await refresh();
    setAgentResult({
      answer: result.success ? result.summary : result.error?.message ?? "Undo is no longer available.",
      actions: [],
      provider: "fallback",
      status: result.success ? "completed" : "error",
    });
  };

  const openSource = (ref: NonNullable<AgentResult["contextRefs"]>[number]) => {
    const routeBySource: Record<NonNullable<AgentResult["contextRefs"]>[number]["type"], RouteId> = {
      file: "learn",
      task: "calendar",
      calendar: "calendar",
      group: "together",
      document: "create",
      memory: "settings",
      settings: "settings",
      focus: "profile",
    };
    navigate(routeBySource[ref.type], { sourceType: ref.type, sourceId: ref.id });
  };
  const compactDock = route === "learn" || route === "create";

  return (
    <div className={`global-agent ${activeConversation ? "is-expanded" : ""} ${compactDock ? "is-compact-dock" : ""}`}>
      {activeConversation && (
        <motion.div className="agent-popover" initial={{ opacity: 0, y: 12 }} animate={{ opacity: 1, y: 0 }}>
          <div className="agent-result-head">
            <span>
              <Bot size={16} />
              LUMA
            </span>
            <small>
              {busy
                ? "checking workspace"
                : agentResult?.provider === "ollama"
                  ? "local model"
                  : agentResult?.status === "offline"
                    ? "local-only mode"
                    : "local operator"}
            </small>
            <button
              aria-label="Close LUMA"
              className="agent-close"
              onClick={() => {
                setExpanded(false);
                setPendingAction(null);
                setAgentResult(null);
              }}
              type="button"
            >
              <X size={14} />
            </button>
          </div>
          {visibleThread.length === 0 && !agentResult && !busy && (
            <div className="agent-command-center">
              <div className="agent-status-card">
                <Orb size="small" />
                <span>
                  <strong>{busy ? "Checking your space..." : "Ask anything. I will use only relevant context."}</strong>
                  <small>{snapshot.settings.localOnlyMaterials ? "Local-only materials are protected" : "Sync context may be available"}</small>
                </span>
              </div>
              <div className="agent-context" aria-label="Current agent context">
                <span>Using</span>
                {quickContext.map((chip) => (
                  <button key={chip} type="button" disabled title="Current context">
                    {chip}
                  </button>
                ))}
              </div>
              <div className="agent-suggestions" aria-label="Try asking">
                <span>Try asking</span>
                {suggestions.map((suggestion) => (
                  <button key={suggestion.label} onClick={() => applySuggestedPrompt(suggestion.prompt)} type="button">
                    {suggestion.label}
                  </button>
                ))}
              </div>
              {recentPrompts.length > 0 && (
                <div className="agent-recent" aria-label="Recent commands">
                  <span>Recent</span>
                  {recentPrompts.map((message) => (
                    <button key={message.id} onClick={() => applySuggestedPrompt(message.content)} type="button">
                      {message.content}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )}
          {(agentResult?.contextChips && agentResult.contextChips.length > 0) || visibleThread.length > 0 ? (
            <div className="agent-context" aria-label="Agent context">
              <span>Using</span>
              {(agentResult?.contextChips ?? quickContext.slice(0, 2)).map((chip) => (
                <button key={chip} type="button" disabled title="Context currently in use">
                  {chip}
                </button>
              ))}
            </div>
          ) : null}
          {visibleThread.length > 0 && (
            <div className="agent-transcript" ref={transcriptRef} aria-label="LUMA chat conversation">
              {visibleThread.map((message) => (
                <article key={message.id} className={`agent-message agent-message-${message.role}`}>
                  <span>{message.role === "user" ? "You" : message.role === "tool" ? "Tool" : "LUMA"}</span>
                  <AgentMessageContent content={message.content} />
                  {message.role === "assistant" && message.contextRefs && message.contextRefs.length > 0 && (
                    <div className="agent-message-sources" aria-label="Message sources">
                      {message.contextRefs.slice(0, 3).map((ref) => (
                        <button key={`${message.id}-${ref.type}-${ref.id}`} type="button" onClick={() => openSource(ref)} title={ref.location ?? `Open ${ref.type}`}>
                          {ref.title}
                        </button>
                      ))}
                    </div>
                  )}
                </article>
              ))}
              {busy && (
                <article className="agent-message agent-message-assistant agent-message-thinking">
                  <span>LUMA</span>
                  <AgentMessageContent content="Checking your workspace..." />
                </article>
              )}
            </div>
          )}
          {visibleThread.length === 0 && agentResult && <p className="agent-answer">{agentResult.answer}</p>}
          {pendingAction && (
            <div className="agent-confirm" role="dialog" aria-label={pendingAction.confirmation?.title ?? "Confirm action"}>
              <small>{pendingAction.requiresConfirmation ? "Confirmation required" : "Plan preview"}</small>
              <strong>{pendingAction.confirmation?.title ?? pendingAction.preview?.title ?? pendingAction.label}</strong>
              <p>{pendingAction.confirmation?.description ?? pendingAction.preview?.summary ?? "Review this action before LUMA changes your workspace."}</p>
              {pendingAction.preview && (
                <ol>
                  {pendingAction.preview.steps.map((step) => (
                    <li key={step}>{step}</li>
                  ))}
                </ol>
              )}
              <div>
                <button
                  type="button"
                  onClick={() => {
                    const action = pendingAction;
                    setPendingAction(null);
                    void executeAction(action);
                  }}
                >
                  {pendingAction.requiresConfirmation ? pendingAction.label : "Run action"}
                </button>
                <button type="button" onClick={() => setPendingAction(null)}>
                  Cancel
                </button>
              </div>
            </div>
          )}
          <div className="agent-actions">
            {agentResult?.actions.map((action) => (
              <button key={action.id} onClick={() => chooseAction(action)} type="button" data-tone={action.style ?? "primary"}>
                {action.label}
                <ChevronRight size={14} />
              </button>
            ))}
            {agentResult?.undo && (
              <button onClick={() => void runUndo()} type="button" data-tone="secondary">
                {agentResult.undo.label}
                <RotateCcw size={14} />
              </button>
            )}
          </div>
        </motion.div>
      )}
      <form
        className="agent-input"
        onClick={() => {
          if (compactDock && !activeConversation) {
            setExpanded(true);
            window.setTimeout(() => inputRef.current?.focus(), 0);
          }
        }}
        onSubmit={(event) => {
          event.preventDefault();
          void submit();
        }}
      >
        <Orb size="small" />
        <input
          ref={inputRef}
          value={input}
          onChange={(event) => setInput(event.currentTarget.value)}
          onFocus={() => setExpanded(true)}
          placeholder={busy ? "LUMA is thinking..." : "Ask LUMA anything..."}
          aria-label="Ask LUMA anything"
        />
        <button disabled={listening} onClick={startVoice} type="button" title={listening ? "Listening..." : "Voice input"}>
          <Mic size={17} />
        </button>
        <button disabled={busy} type="submit" title="Send">
          <Send size={17} />
        </button>
      </form>
    </div>
  );
}
