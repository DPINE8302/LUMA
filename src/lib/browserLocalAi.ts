import type { ChatCompletionMessageParam } from "@mlc-ai/web-llm/lib/openai_api_protocols";
import type { AgentMessage, AgentResult, LumaSnapshot } from "./types";

export interface BrowserLocalModel {
  id: string;
  name: string;
  description: string;
  sizeLabel: string;
  downloadSize: string;
  vramRequired: string;
  isDefault?: boolean;
}

export type BrowserLocalStatus = "unsupported" | "not-loaded" | "loading" | "ready" | "generating" | "error";

export interface BrowserLocalSupport {
  supported: boolean;
  webgpu: boolean;
  reason?: string;
}

export interface BrowserLocalProgress {
  progress: number;
  message: string;
}

export const BROWSER_LOCAL_MODELS: BrowserLocalModel[] = [
  {
    id: "Llama-3.2-1B-Instruct-q4f16_1-MLC",
    name: "Llama 3.2 1B Instruct",
    description: "Fast Meta model for local student planning and short answers.",
    sizeLabel: "Ultra-Lightweight",
    downloadSize: "~610 MB",
    vramRequired: "~1.2 GB",
    isDefault: true,
  },
  {
    id: "Qwen2.5-1.5B-Instruct-q4f16_1-MLC",
    name: "Qwen 2.5 1.5B Instruct",
    description: "Stronger multilingual reasoning with a larger download.",
    sizeLabel: "Lightweight",
    downloadSize: "~1.1 GB",
    vramRequired: "~1.8 GB",
  },
  {
    id: "SmolLM2-360M-Instruct-q4f16_1-MLC",
    name: "SmolLM2 360M Instruct",
    description: "Tiny fallback model for older or memory-limited devices.",
    sizeLabel: "Tiny",
    downloadSize: "~250 MB",
    vramRequired: "<1 GB",
  },
];

export const DEFAULT_BROWSER_LOCAL_MODEL = BROWSER_LOCAL_MODELS.find((model) => model.isDefault) ?? BROWSER_LOCAL_MODELS[0];

let engine: import("@mlc-ai/web-llm").MLCEngine | null = null;
let loadedModelId: string | null = null;
let loadingPromise: Promise<void> | null = null;

function browserAvailable() {
  return typeof window !== "undefined" && typeof navigator !== "undefined";
}

function getErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : String(error);
}

export async function checkBrowserLocalSupport(): Promise<BrowserLocalSupport> {
  if (!browserAvailable()) {
    return { supported: false, webgpu: false, reason: "Browser local AI needs a browser runtime." };
  }
  if (!("gpu" in navigator)) {
    return {
      supported: false,
      webgpu: false,
      reason: "WebGPU is not available in this browser. Use Chrome, Edge, or another WebGPU-enabled desktop browser.",
    };
  }
  try {
    const gpuNavigator = navigator as Navigator & { gpu?: { requestAdapter: () => Promise<unknown> } };
    const adapter = await gpuNavigator.gpu?.requestAdapter();
    return adapter
      ? { supported: true, webgpu: true }
      : { supported: false, webgpu: true, reason: "WebGPU is available, but no compatible GPU adapter was found." };
  } catch (error) {
    return { supported: false, webgpu: true, reason: `WebGPU check failed: ${getErrorMessage(error)}` };
  }
}

export async function checkBrowserModelCached(modelId = DEFAULT_BROWSER_LOCAL_MODEL.id): Promise<boolean> {
  if (!browserAvailable()) return false;
  try {
    const { hasModelInCache } = await import("@mlc-ai/web-llm");
    return await hasModelInCache(modelId);
  } catch {
    return false;
  }
}

export async function deleteBrowserModelCache(modelId = DEFAULT_BROWSER_LOCAL_MODEL.id): Promise<void> {
  const { deleteModelAllInfoInCache } = await import("@mlc-ai/web-llm");
  if (loadedModelId === modelId) {
    await unloadBrowserLocalModel();
  }
  await deleteModelAllInfoInCache(modelId);
}

export async function unloadBrowserLocalModel() {
  if (engine) {
    await engine.unload();
  }
  engine = null;
  loadedModelId = null;
  loadingPromise = null;
}

export function getBrowserLocalRuntimeState() {
  return {
    ready: Boolean(engine && loadedModelId),
    modelId: loadedModelId,
  };
}

export async function initializeBrowserLocalModel(
  modelId = DEFAULT_BROWSER_LOCAL_MODEL.id,
  onProgress?: (progress: BrowserLocalProgress) => void,
) {
  if (engine && loadedModelId === modelId) return;
  if (loadingPromise) return loadingPromise;

  loadingPromise = (async () => {
    const support = await checkBrowserLocalSupport();
    if (!support.supported) {
      throw new Error(support.reason ?? "Browser local AI is not supported here.");
    }
    await unloadBrowserLocalModel();
    const { CreateMLCEngine } = await import("@mlc-ai/web-llm");
    engine = await CreateMLCEngine(modelId, {
      initProgressCallback: (progress) => {
        onProgress?.({ progress: progress.progress, message: progress.text });
      },
    });
    loadedModelId = modelId;
  })();

  try {
    await loadingPromise;
  } finally {
    loadingPromise = null;
  }
}

export async function generateBrowserLocalAgentResult(
  input: string,
  snapshot: LumaSnapshot,
  modelId = DEFAULT_BROWSER_LOCAL_MODEL.id,
  options: { scope?: "personal" | "group"; recentMessages?: AgentMessage[] } = {},
): Promise<AgentResult> {
  if (!engine || loadedModelId !== modelId) {
    await initializeBrowserLocalModel(modelId);
  }
  if (!engine) {
    throw new Error("Browser local AI did not initialize.");
  }

  const scope = options.scope ?? "personal";
  const group = snapshot.groups[0];
  const scopedTasks = scope === "group" && group
    ? snapshot.tasks.filter((task) => group.taskIds.includes(task.id) || task.groupId === group.id)
    : snapshot.tasks;
  const scopedMaterials = scope === "group" && group
    ? snapshot.materials.filter((material) => material.sharedWith.includes(group.id) || group.materialIds.includes(material.id))
    : snapshot.materials;
  const context = scopedTasks
    .filter((task) => task.status !== "done")
    .slice(0, 6)
    .map((task) => `${task.title} due ${task.dueAt}; notes: ${task.notes}`)
    .join("\n");
  const materials = scopedMaterials
    .slice(0, 4)
    .map((material) => `${material.title} (${material.folder}): ${material.content.slice(0, 600)}`)
    .join("\n");
  const groups = scope === "group" && group
    ? `${group.name}: ${group.progress}% progress; milestones ${group.milestones.map((item) => `${item.label}:${item.done ? "done" : "open"}`).join(", ")}`
    : "Personal mode; do not use shared group context unless the user asks for it.";
  const recentChat = (options.recentMessages ?? [])
    .filter((message) => message.role === "user" || message.role === "assistant")
    .slice(-6)
    .map((message) => ({
      role: message.role,
      content: message.content,
    })) as ChatCompletionMessageParam[];

  const messages: ChatCompletionMessageParam[] = [
    {
      role: "system",
      content:
        `You are LUMA Agent, a calm local-first school operator. Answer concisely with clean paragraph and bullet formatting when useful. Use only supplied ${scope} context, name sources when relevant, and suggest one useful next action. Do not claim to change data unless a tool is shown.`,
    },
    ...recentChat,
    {
      role: "user",
      content: `Current LUMA ${scope} context:\nTasks:\n${context || "None"}\nMaterials:\n${materials || "None"}\nGroups:\n${groups || "None"}\n\nUser request: ${input}`,
    },
  ];

  const reply = await engine.chat.completions.create({
    messages,
    stream: false,
    temperature: 0.2,
    max_tokens: 260,
  });

  return {
    answer: reply.choices[0]?.message.content?.trim() || "I could not produce a local model answer.",
    actions: [],
    provider: "local",
  };
}
