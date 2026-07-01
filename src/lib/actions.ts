import { addDays } from "date-fns";
import { getCurrentOwnerId, lumaDb } from "./db";
import { gradeFlashcard, moveTaskByNaturalLanguage } from "./scheduling";
import type {
  ActionHistoryEntry,
  AgentToolResult,
  AuditImpact,
  AuditLogEntry,
  ChatMessage,
  ChecklistItem,
  CreateDocument,
  DocumentVersion,
  DocumentType,
  EntityMeta,
  Flashcard,
  FlashcardReview,
  GroupWorkspace,
  Highlight,
  Material,
  Quiz,
  QuizAttempt,
  SourceReference,
  StudySession,
  TaskItem,
  TaskStatus,
  TaskType,
} from "./types";

const uid = (prefix: string) => `${prefix}-${crypto.randomUUID()}`;

const ownerId = () => getCurrentOwnerId();

export function createEntityMeta(owner = ownerId(), now = new Date().toISOString()): EntityMeta {
  return {
    ownerId: owner,
    createdAt: now,
    updatedAt: now,
    syncState: "local",
    permissionState: "owner",
    visibility: "private",
  };
}

function asTaskRef(task: TaskItem): SourceReference {
  return {
    type: "task",
    id: task.id,
    title: task.title,
    location: task.dueAt,
  };
}

function asDocumentRef(document: CreateDocument): SourceReference {
  return {
    type: "document",
    id: document.id,
    title: document.title,
    location: document.status,
  };
}

function versionFromDocument(document: CreateDocument, label: string, now = new Date().toISOString()): DocumentVersion {
  return {
    id: uid("doc-version"),
    ownerId: ownerId(),
    documentId: document.id,
    title: document.title,
    body: document.body,
    outline: document.outline,
    label,
    createdAt: now,
    syncState: "local",
    permissionState: "owner",
  };
}

function asMaterialRef(material: Material): SourceReference {
  return {
    type: "file",
    id: material.id,
    title: material.title,
    location: material.folder,
  };
}

function asFlashcardRef(card: Flashcard): SourceReference {
  return {
    type: "file",
    id: card.id,
    title: card.front,
    location: card.deck,
  };
}

function asQuizRef(quiz: Quiz): SourceReference {
  return {
    type: "file",
    id: quiz.id,
    title: quiz.title,
    location: "Quiz",
  };
}

async function recordAction({
  action,
  summary,
  impact,
  sourceRefs = [],
  affected,
  payload,
  inversePayload,
  provider = "manual",
}: {
  action: string;
  summary: string;
  impact: AuditImpact;
  sourceRefs?: SourceReference[];
  affected?: SourceReference;
  payload: Record<string, unknown>;
  inversePayload?: Record<string, unknown>;
  provider?: AuditLogEntry["provider"];
}) {
  const now = new Date().toISOString();
  const actionId = uid("action");
  const history: ActionHistoryEntry = {
    id: actionId,
    ownerId: ownerId(),
    action,
    status: "completed",
    summary,
    payload,
    inversePayload,
    createdAt: now,
    updatedAt: now,
    expiresAt: inversePayload ? addDays(new Date(now), 1).toISOString() : undefined,
  };
  const audit: AuditLogEntry = {
    id: uid("audit"),
    ownerId: ownerId(),
    action,
    summary,
    impact,
    sourceRefs,
    affectedType: affected?.type,
    affectedId: affected?.id,
    provider,
    createdAt: now,
    undoActionId: inversePayload ? actionId : undefined,
  };
  await lumaDb.transaction("rw", lumaDb.actionHistory, lumaDb.auditLog, async () => {
    await lumaDb.actionHistory.put(history);
    await lumaDb.auditLog.put(audit);
  });
  return history;
}

export async function createGroupTool(input: {
  name: string;
  subjectId: string;
}): Promise<AgentToolResult<GroupWorkspace>> {
  const name = input.name.trim();
  if (!name) {
    return {
      success: false,
      summary: "Group was not created.",
      error: {
        code: "empty_group_name",
        message: "A workspace needs a name before LUMA can create it.",
        recoveryActions: ["Enter a group name"],
      },
    };
  }
  const group: GroupWorkspace = {
    id: uid("group"),
    ownerId: ownerId(),
    name,
    subjectId: input.subjectId,
    members: [
      {
        id: ownerId(),
        name: "You",
        role: "owner",
        avatar: "You",
        online: true,
      },
    ],
    taskIds: [],
    materialIds: [],
    progress: 0,
    milestones: [],
  };
  const groupRef: SourceReference = { type: "group", id: group.id, title: group.name };
  const history = await recordAction({
    action: "create_group",
    summary: `Created group workspace "${group.name}".`,
    impact: "personal",
    affected: groupRef,
    payload: { group },
    inversePayload: { deleteGroupId: group.id },
  });
  await lumaDb.groups.put(group);
  return {
    success: true,
    data: group,
    summary: `Created group workspace "${group.name}".`,
    sourceRefs: [groupRef],
    undo: {
      label: "Undo group creation",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function postGroupMessageTool(input: {
  groupId: string;
  message: string;
}): Promise<AgentToolResult<ChatMessage>> {
  const group = await lumaDb.groups.get(input.groupId);
  const message = input.message.trim();
  if (!group) {
    return {
      success: false,
      summary: "Message was not sent.",
      error: {
        code: "group_not_found",
        message: "LUMA could not find that group workspace.",
        recoveryActions: ["Create a group"],
      },
    };
  }
  if (!message) {
    return {
      success: false,
      summary: "Message was not sent.",
      error: {
        code: "empty_message",
        message: "A group message cannot be empty.",
        recoveryActions: ["Write a message"],
      },
    };
  }
  const chat: ChatMessage = {
    id: uid("chat"),
    ownerId: ownerId(),
    groupId: group.id,
    author: "You",
    message,
    createdAt: new Date().toISOString(),
  };
  const chatRef: SourceReference = { type: "group", id: group.id, title: group.name, location: "Group chat" };
  const history = await recordAction({
    action: "post_group_message",
    summary: `Posted a message to "${group.name}".`,
    impact: "shared",
    sourceRefs: [chatRef],
    affected: chatRef,
    payload: { chat },
    inversePayload: { deleteChatId: chat.id },
  });
  await lumaDb.chats.put(chat);
  return {
    success: true,
    data: chat,
    summary: `Posted a message to "${group.name}".`,
    sourceRefs: [chatRef],
    undo: {
      label: "Undo message",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateGroupMessageTool(input: {
  chatId: string;
  message: string;
}): Promise<AgentToolResult<ChatMessage>> {
  const chat = await lumaDb.chats.get(input.chatId);
  const message = input.message.trim();
  if (!chat || !message) {
    return {
      success: false,
      summary: "Message was not updated.",
      error: {
        code: !chat ? "chat_not_found" : "empty_message",
        message: !chat ? "LUMA could not find that chat message." : "A group message cannot be empty.",
        recoveryActions: ["Open group chat"],
      },
    };
  }
  if (chat.author !== "You") {
    return {
      success: false,
      summary: "Message was not updated.",
      error: {
        code: "permission_denied",
        message: "Only your own messages can be edited locally.",
        recoveryActions: ["Ask the author to edit their message"],
      },
    };
  }
  const updated: ChatMessage = { ...chat, message };
  const chatRef: SourceReference = { type: "group", id: chat.groupId, title: "Group chat" };
  const history = await recordAction({
    action: "update_group_message",
    summary: "Edited a group message.",
    impact: "shared",
    sourceRefs: [chatRef],
    affected: chatRef,
    payload: { chatId: chat.id, message },
    inversePayload: { chat },
  });
  await lumaDb.chats.put(updated);
  return {
    success: true,
    data: updated,
    summary: "Edited a group message.",
    sourceRefs: [chatRef],
    undo: {
      label: "Undo message edit",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function deleteGroupMessageTool(chatId: string): Promise<AgentToolResult> {
  const chat = await lumaDb.chats.get(chatId);
  if (!chat) {
    return {
      success: false,
      summary: "Message was not deleted.",
      error: {
        code: "chat_not_found",
        message: "LUMA could not find that chat message.",
        recoveryActions: ["Open group chat"],
      },
    };
  }
  if (chat.author !== "You") {
    return {
      success: false,
      summary: "Message was not deleted.",
      error: {
        code: "permission_denied",
        message: "Only your own messages can be deleted locally.",
        recoveryActions: ["Ask the author to delete their message"],
      },
    };
  }
  const chatRef: SourceReference = { type: "group", id: chat.groupId, title: "Group chat" };
  const history = await recordAction({
    action: "delete_group_message",
    summary: "Deleted a group message.",
    impact: "shared",
    sourceRefs: [chatRef],
    affected: chatRef,
    payload: { chatId },
    inversePayload: { chat },
  });
  await lumaDb.chats.delete(chatId);
  return {
    success: true,
    summary: "Deleted a group message.",
    sourceRefs: [chatRef],
    undo: {
      label: "Undo message delete",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Delete message",
      description: "This removes your message from the group chat.",
      impact: "shared",
    },
  };
}

export async function shareMaterialToGroupTool(input: {
  materialId: string;
  groupId: string;
}): Promise<AgentToolResult<Material>> {
  const [material, group] = await Promise.all([lumaDb.materials.get(input.materialId), lumaDb.groups.get(input.groupId)]);
  if (!material || !group) {
    return {
      success: false,
      summary: "Material was not shared.",
      error: {
        code: "share_target_missing",
        message: "LUMA could not find both the material and the group workspace.",
        recoveryActions: ["Open Study Vault", "Open Together"],
      },
    };
  }
  const updated: Material = {
    ...material,
    sharedWith: [...new Set([...material.sharedWith, group.id])],
    updatedAt: new Date().toISOString(),
    meta: {
      ...(material.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local",
      visibility: "group",
    },
  };
  const materialRef: SourceReference = { type: "file", id: material.id, title: material.title };
  const groupRef: SourceReference = { type: "group", id: group.id, title: group.name };
  const history = await recordAction({
    action: "share_material_to_group",
    summary: `Shared "${material.title}" with "${group.name}".`,
    impact: "shared",
    sourceRefs: [materialRef, groupRef],
    affected: materialRef,
    payload: { materialId: material.id, groupId: group.id },
    inversePayload: { materialId: material.id, sharedWith: material.sharedWith },
  });
  await lumaDb.materials.put(updated);
  return {
    success: true,
    data: updated,
    summary: `Shared "${material.title}" with "${group.name}".`,
    sourceRefs: [materialRef, groupRef],
    undo: {
      label: "Undo share",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Share material",
      description: `This shares "${material.title}" with "${group.name}".`,
      impact: "shared",
    },
  };
}

export async function unshareMaterialFromGroupTool(input: {
  materialId: string;
  groupId: string;
}): Promise<AgentToolResult<Material>> {
  const [material, group] = await Promise.all([lumaDb.materials.get(input.materialId), lumaDb.groups.get(input.groupId)]);
  if (!material || !group) {
    return {
      success: false,
      summary: "Material was not removed from the group.",
      error: {
        code: "share_target_missing",
        message: "LUMA could not find both the material and the group workspace.",
        recoveryActions: ["Open Study Vault", "Open Together"],
      },
    };
  }
  const updated: Material = {
    ...material,
    sharedWith: material.sharedWith.filter((id) => id !== group.id),
    updatedAt: new Date().toISOString(),
    meta: {
      ...(material.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local",
      visibility: material.sharedWith.length > 1 ? "group" : "private",
    },
  };
  const materialRef = asMaterialRef(material);
  const groupRef: SourceReference = { type: "group", id: group.id, title: group.name };
  const history = await recordAction({
    action: "unshare_material_from_group",
    summary: `Removed "${material.title}" from "${group.name}".`,
    impact: "shared",
    sourceRefs: [materialRef, groupRef],
    affected: materialRef,
    payload: { materialId: material.id, groupId: group.id },
    inversePayload: { materialId: material.id, sharedWith: material.sharedWith },
  });
  await lumaDb.materials.put(updated);
  return {
    success: true,
    data: updated,
    summary: `Removed "${material.title}" from "${group.name}".`,
    sourceRefs: [materialRef, groupRef],
    undo: {
      label: "Undo file removal",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Remove shared file",
      description: `This removes "${material.title}" from "${group.name}".`,
      impact: "shared",
    },
  };
}

export async function updateGroupMemberRoleTool(input: {
  groupId: string;
  memberId: string;
  role: string;
}): Promise<AgentToolResult<GroupWorkspace>> {
  const group = await lumaDb.groups.get(input.groupId);
  if (!group) {
    return {
      success: false,
      summary: "Member role was not updated.",
      error: {
        code: "group_not_found",
        message: "LUMA could not find that group workspace.",
        recoveryActions: ["Open Together"],
      },
    };
  }
  const member = group.members.find((item) => item.id === input.memberId);
  if (!member) {
    return {
      success: false,
      summary: "Member role was not updated.",
      error: {
        code: "member_not_found",
        message: "LUMA could not find that group member.",
        recoveryActions: ["Open members"],
      },
    };
  }
  const updated: GroupWorkspace = {
    ...group,
    members: group.members.map((item) => (item.id === input.memberId ? { ...item, role: input.role } : item)),
  };
  const groupRef: SourceReference = { type: "group", id: group.id, title: group.name };
  const history = await recordAction({
    action: "update_group_member_role",
    summary: `Changed ${member.name}'s role to ${input.role}.`,
    impact: "shared",
    sourceRefs: [groupRef],
    affected: groupRef,
    payload: { groupId: group.id, memberId: member.id, role: input.role },
    inversePayload: { group },
  });
  await lumaDb.groups.put(updated);
  return {
    success: true,
    data: updated,
    summary: `Changed ${member.name}'s role to ${input.role}.`,
    sourceRefs: [groupRef],
    undo: {
      label: "Undo role change",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Change member role",
      description: `This changes ${member.name}'s role in "${group.name}".`,
      impact: "shared",
    },
  };
}

export async function createMaterialTool(material: Material): Promise<AgentToolResult<Material>> {
  const now = new Date().toISOString();
  const saved: Material = {
    ...material,
    createdAt: material.createdAt || now,
    updatedAt: now,
    meta: material.meta ?? createEntityMeta(ownerId(), now),
  };
  const history = await recordAction({
    action: "create_material",
    summary: `Added "${saved.title}" to Study Vault.`,
    impact: "personal",
    affected: asMaterialRef(saved),
    payload: { material: saved },
    inversePayload: { deleteMaterialId: saved.id },
  });
  await lumaDb.materials.put(saved);
  return {
    success: true,
    data: saved,
    summary: `Added "${saved.title}" to Study Vault.`,
    sourceRefs: [asMaterialRef(saved)],
    undo: {
      label: "Undo material upload",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateMaterialTool(input: {
  materialId: string;
  patch: Partial<Pick<Material, "title" | "subjectId" | "folder" | "tags" | "content">>;
  summary?: string;
}): Promise<AgentToolResult<Material>> {
  const material = await lumaDb.materials.get(input.materialId);
  if (!material) {
    return {
      success: false,
      summary: "Material was not updated.",
      error: {
        code: "material_not_found",
        message: "LUMA could not find that Study Vault material.",
        recoveryActions: ["Open Study Vault"],
      },
    };
  }
  const updated: Material = {
    ...material,
    ...input.patch,
    updatedAt: new Date().toISOString(),
    meta: {
      ...(material.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local",
    },
  };
  const history = await recordAction({
    action: "update_material",
    summary: input.summary ?? `Updated "${updated.title}".`,
    impact: "personal",
    sourceRefs: [asMaterialRef(material)],
    affected: asMaterialRef(updated),
    payload: { materialId: material.id, patch: input.patch },
    inversePayload: { material },
  });
  await lumaDb.materials.put(updated);
  return {
    success: true,
    data: updated,
    summary: input.summary ?? `Updated "${updated.title}".`,
    sourceRefs: [asMaterialRef(updated)],
    undo: {
      label: "Undo material update",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function deleteMaterialTool(materialId: string): Promise<AgentToolResult> {
  const material = await lumaDb.materials.get(materialId);
  if (!material) {
    return {
      success: false,
      summary: "Material was not deleted.",
      error: {
        code: "material_not_found",
        message: "LUMA could not find that Study Vault material.",
        recoveryActions: ["Open Study Vault"],
      },
    };
  }
  const highlights = await lumaDb.highlights.where("materialId").equals(material.id).toArray();
  const history = await recordAction({
    action: "delete_material",
    summary: `Deleted "${material.title}" from Study Vault.`,
    impact: "destructive",
    sourceRefs: [asMaterialRef(material)],
    affected: asMaterialRef(material),
    payload: { materialId },
    inversePayload: { material, highlights },
  });
  await lumaDb.transaction("rw", lumaDb.materials, lumaDb.highlights, async () => {
    await lumaDb.materials.delete(material.id);
    await Promise.all(highlights.map((item) => lumaDb.highlights.delete(item.id)));
  });
  return {
    success: true,
    summary: `Deleted "${material.title}" from Study Vault.`,
    sourceRefs: [asMaterialRef(material)],
    undo: {
      label: "Undo material delete",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Delete material",
      description: `This deletes "${material.title}" and its highlights.`,
      impact: "destructive",
    },
  };
}

export async function createHighlightTool(input: {
  materialId: string;
  text: string;
  note?: string;
}): Promise<AgentToolResult<Highlight>> {
  const material = await lumaDb.materials.get(input.materialId);
  if (!material) {
    return {
      success: false,
      summary: "Highlight was not saved.",
      error: {
        code: "material_not_found",
        message: "LUMA could not find that Study Vault material.",
        recoveryActions: ["Open Study Vault"],
      },
    };
  }
  const highlight: Highlight = {
    id: uid("highlight"),
    ownerId: ownerId(),
    materialId: material.id,
    text: input.text.trim() || material.content.slice(0, 120),
    note: input.note ?? "Saved from reader",
    createdAt: new Date().toISOString(),
  };
  const history = await recordAction({
    action: "create_highlight",
    summary: `Saved a highlight from "${material.title}".`,
    impact: "personal",
    sourceRefs: [asMaterialRef(material)],
    affected: asMaterialRef(material),
    payload: { highlight },
    inversePayload: { deleteHighlightId: highlight.id },
  });
  await lumaDb.highlights.put(highlight);
  return {
    success: true,
    data: highlight,
    summary: `Saved a highlight from "${material.title}".`,
    sourceRefs: [asMaterialRef(material)],
    undo: {
      label: "Undo highlight",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function createFlashcardsTool(input: {
  materialId: string;
  flashcards: Flashcard[];
}): Promise<AgentToolResult<Flashcard[]>> {
  const material = await lumaDb.materials.get(input.materialId);
  if (!material) {
    return {
      success: false,
      summary: "Flashcards were not created.",
      error: {
        code: "material_not_found",
        message: "LUMA could not find that Study Vault material.",
        recoveryActions: ["Open Study Vault"],
      },
    };
  }
  const cards = input.flashcards.map((card) => ({
    ...card,
    ownerId: card.ownerId ?? ownerId(),
    sourceMaterialId: material.id,
  }));
  const history = await recordAction({
    action: "create_flashcards",
    summary: `Created ${cards.length} flashcard${cards.length === 1 ? "" : "s"} from "${material.title}".`,
    impact: "personal",
    sourceRefs: [asMaterialRef(material)],
    affected: asMaterialRef(material),
    payload: { materialId: material.id, flashcardIds: cards.map((card) => card.id) },
    inversePayload: { deleteFlashcardIds: cards.map((card) => card.id) },
  });
  await lumaDb.flashcards.bulkPut(cards);
  return {
    success: true,
    data: cards,
    summary: `Created ${cards.length} flashcard${cards.length === 1 ? "" : "s"} from "${material.title}".`,
    sourceRefs: [asMaterialRef(material)],
    undo: {
      label: "Undo flashcards",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function createQuizTool(input: {
  materialId: string;
  quiz: Quiz;
}): Promise<AgentToolResult<Quiz>> {
  const material = await lumaDb.materials.get(input.materialId);
  if (!material) {
    return {
      success: false,
      summary: "Quiz was not created.",
      error: {
        code: "material_not_found",
        message: "LUMA could not find that Study Vault material.",
        recoveryActions: ["Open Study Vault"],
      },
    };
  }
  const quiz = { ...input.quiz, ownerId: input.quiz.ownerId ?? ownerId() };
  const history = await recordAction({
    action: "create_quiz",
    summary: `Created quiz "${quiz.title}" from "${material.title}".`,
    impact: "personal",
    sourceRefs: [asMaterialRef(material)],
    affected: asMaterialRef(material),
    payload: { materialId: material.id, quizId: quiz.id },
    inversePayload: { deleteQuizId: quiz.id },
  });
  await lumaDb.quizzes.put(quiz);
  return {
    success: true,
    data: quiz,
    summary: `Created quiz "${quiz.title}" from "${material.title}".`,
    sourceRefs: [asMaterialRef(material)],
    undo: {
      label: "Undo quiz",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function reviewFlashcardTool(input: {
  flashcardId: string;
  quality: FlashcardReview["quality"];
}): Promise<AgentToolResult<FlashcardReview>> {
  const card = await lumaDb.flashcards.get(input.flashcardId);
  if (!card) {
    return {
      success: false,
      summary: "Flashcard review was not saved.",
      error: {
        code: "flashcard_not_found",
        message: "LUMA could not find that flashcard.",
        recoveryActions: ["Open Learn"],
      },
    };
  }
  const reviewed = gradeFlashcard(card, input.quality);
  const now = new Date().toISOString();
  const review: FlashcardReview = {
    id: uid("review"),
    ownerId: ownerId(),
    flashcardId: card.id,
    subjectId: card.subjectId,
    quality: input.quality,
    previousMastery: card.mastery,
    nextMastery: reviewed.mastery,
    reviewedAt: now,
    nextDueAt: reviewed.dueAt,
    syncState: "local",
    permissionState: "owner",
  };
  const history = await recordAction({
    action: "review_flashcard",
    summary: `Reviewed flashcard "${card.front}".`,
    impact: "personal",
    sourceRefs: [asFlashcardRef(card)],
    affected: asFlashcardRef(card),
    payload: { review, flashcard: reviewed },
    inversePayload: { flashcard: card, deleteFlashcardReviewId: review.id },
  });
  await lumaDb.transaction("rw", lumaDb.flashcards, lumaDb.flashcardReviews, async () => {
    await lumaDb.flashcards.put(reviewed);
    await lumaDb.flashcardReviews.put(review);
  });
  return {
    success: true,
    data: review,
    summary: `Reviewed "${card.front}" as ${input.quality}.`,
    sourceRefs: [asFlashcardRef(reviewed)],
    undo: {
      label: "Undo review",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function submitQuizAttemptTool(input: {
  quizId: string;
  answers: Record<string, string>;
}): Promise<AgentToolResult<QuizAttempt>> {
  const quiz = await lumaDb.quizzes.get(input.quizId);
  if (!quiz) {
    return {
      success: false,
      summary: "Quiz attempt was not saved.",
      error: {
        code: "quiz_not_found",
        message: "LUMA could not find that quiz.",
        recoveryActions: ["Open Learn"],
      },
    };
  }
  const correctQuestions = quiz.questions.filter((question) => {
    const answer = input.answers[question.id]?.trim().toLowerCase();
    return Boolean(answer) && answer === question.answer.trim().toLowerCase();
  });
  const weakTopics = quiz.questions
    .filter((question) => !correctQuestions.some((correct) => correct.id === question.id))
    .map((question) => question.topic)
    .filter((topic, index, topics) => topics.indexOf(topic) === index);
  const score = quiz.questions.length ? Math.round((correctQuestions.length / quiz.questions.length) * 100) : 0;
  const submittedAt = new Date().toISOString();
  const attempt: QuizAttempt = {
    id: uid("attempt"),
    ownerId: ownerId(),
    quizId: quiz.id,
    subjectId: quiz.subjectId,
    answers: input.answers,
    score,
    totalQuestions: quiz.questions.length,
    correctCount: correctQuestions.length,
    weakTopics,
    startedAt: submittedAt,
    submittedAt,
    syncState: "local",
    permissionState: "owner",
  };
  const updatedQuiz: Quiz = {
    ...quiz,
    score,
    weakTopics,
    completedAt: submittedAt,
  };
  const history = await recordAction({
    action: "submit_quiz_attempt",
    summary: `Submitted "${quiz.title}" with ${score}%.`,
    impact: "personal",
    sourceRefs: [asQuizRef(quiz)],
    affected: asQuizRef(quiz),
    payload: { attempt, quiz: updatedQuiz },
    inversePayload: { quiz, deleteQuizAttemptId: attempt.id },
  });
  await lumaDb.transaction("rw", lumaDb.quizzes, lumaDb.quizAttempts, async () => {
    await lumaDb.quizzes.put(updatedQuiz);
    await lumaDb.quizAttempts.put(attempt);
  });
  return {
    success: true,
    data: attempt,
    summary: `Submitted "${quiz.title}" with ${score}%.`,
    sourceRefs: [asQuizRef(updatedQuiz)],
    undo: {
      label: "Undo quiz submission",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function createTaskTool(input: {
  title: string;
  subjectId: string;
  dueAt: string;
  priority?: TaskItem["priority"];
  type?: TaskType;
  notes?: string;
  groupId?: string;
  assignee?: string;
}): Promise<AgentToolResult<TaskItem>> {
  const title = input.title.trim();
  if (!title) {
    return {
      success: false,
      summary: "Task was not created.",
      error: {
        code: "empty_task_title",
        message: "A task needs a title before LUMA can save it.",
        recoveryActions: ["Enter a task title"],
      },
    };
  }
  const now = new Date().toISOString();
  const task: TaskItem = {
    id: uid("task"),
    title,
    subjectId: input.subjectId,
    status: "todo",
    dueAt: input.dueAt,
    priority: input.priority ?? "medium",
    type: input.type ?? "homework",
    notes: input.notes ?? "Created in LUMA.",
    groupId: input.groupId,
    assignee: input.assignee,
    checklistItemIds: [],
    meta: createEntityMeta(ownerId(), now),
  };
  const history = await recordAction({
    action: "create_task",
    summary: `Created task "${task.title}".`,
    impact: input.groupId ? "shared" : "personal",
    affected: asTaskRef(task),
    payload: { task },
    inversePayload: { deleteTaskId: task.id },
  });
  await lumaDb.tasks.put(task);
  return {
    success: true,
    data: task,
    summary: `Created task "${task.title}".`,
    sourceRefs: [asTaskRef(task)],
    undo: {
      label: "Undo task creation",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function createDocumentTool(input: {
  title: string;
  type: DocumentType;
  outline: string[];
  body: string;
  dueAt?: string;
  sourceMaterialIds?: string[];
  milestones?: string[];
}): Promise<AgentToolResult<CreateDocument>> {
  const title = input.title.trim() || "Untitled Project";
  const document: CreateDocument = {
    id: uid("doc"),
    ownerId: ownerId(),
    type: input.type,
    title,
    outline: input.outline,
    body: input.body,
    status: "draft",
    dueAt: input.dueAt,
    sourceMaterialIds: input.sourceMaterialIds ?? [],
    milestones: input.milestones ?? ["Details", "Sources", "Outline", "Draft"],
  };
  const history = await recordAction({
    action: "create_document",
    summary: `Created project "${document.title}".`,
    impact: "personal",
    affected: asDocumentRef(document),
    payload: { document },
    inversePayload: { deleteDocumentId: document.id },
  });
  await lumaDb.documents.put(document);
  return {
    success: true,
    data: document,
    summary: `Created project "${document.title}".`,
    sourceRefs: [asDocumentRef(document)],
    undo: {
      label: "Undo project creation",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateDocumentTool(input: {
  documentId: string;
  patch: Partial<Pick<CreateDocument, "title" | "body" | "outline" | "status" | "dueAt" | "milestones" | "sourceMaterialIds">>;
  summary?: string;
  versionLabel?: string;
}): Promise<AgentToolResult<CreateDocument>> {
  const document = await lumaDb.documents.get(input.documentId);
  if (!document) {
    return {
      success: false,
      summary: "Project was not updated.",
      error: {
        code: "document_not_found",
        message: "LUMA could not find that Create project.",
        recoveryActions: ["Open Create"],
      },
    };
  }
  const updated: CreateDocument = { ...document, ...input.patch };
  const version = versionFromDocument(document, input.versionLabel ?? "Before edit");
  const history = await recordAction({
    action: "update_document",
    summary: input.summary ?? `Updated project "${updated.title}".`,
    impact: "personal",
    sourceRefs: [asDocumentRef(document)],
    affected: asDocumentRef(updated),
    payload: { documentId: document.id, patch: input.patch },
    inversePayload: { document, deleteDocumentVersionId: version.id },
  });
  await lumaDb.transaction("rw", lumaDb.documents, lumaDb.documentVersions, async () => {
    await lumaDb.documentVersions.put(version);
    await lumaDb.documents.put(updated);
  });
  return {
    success: true,
    data: updated,
    summary: input.summary ?? `Updated project "${updated.title}".`,
    sourceRefs: [asDocumentRef(updated)],
    undo: {
      label: "Undo project update",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function restoreDocumentVersionTool(versionId: string): Promise<AgentToolResult<CreateDocument>> {
  const version = await lumaDb.documentVersions.get(versionId);
  if (!version) {
    return {
      success: false,
      summary: "Version was not restored.",
      error: {
        code: "document_version_not_found",
        message: "LUMA could not find that document version.",
        recoveryActions: ["Open Create"],
      },
    };
  }
  const document = await lumaDb.documents.get(version.documentId);
  if (!document) {
    return {
      success: false,
      summary: "Version was not restored.",
      error: {
        code: "document_not_found",
        message: "LUMA could not find the project for that version.",
        recoveryActions: ["Open Create"],
      },
    };
  }
  const beforeRestore = versionFromDocument(document, "Before restore");
  const restored: CreateDocument = {
    ...document,
    title: version.title,
    body: version.body,
    outline: version.outline,
  };
  const history = await recordAction({
    action: "restore_document_version",
    summary: `Restored "${document.title}" to ${version.label}.`,
    impact: "personal",
    sourceRefs: [asDocumentRef(document)],
    affected: asDocumentRef(restored),
    payload: { versionId, documentId: document.id },
    inversePayload: { document, deleteDocumentVersionId: beforeRestore.id },
  });
  await lumaDb.transaction("rw", lumaDb.documents, lumaDb.documentVersions, async () => {
    await lumaDb.documentVersions.put(beforeRestore);
    await lumaDb.documents.put(restored);
  });
  return {
    success: true,
    data: restored,
    summary: `Restored "${document.title}" to ${version.label}.`,
    sourceRefs: [asDocumentRef(restored)],
    undo: {
      label: "Undo restore",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function duplicateDocumentTool(documentId: string): Promise<AgentToolResult<CreateDocument>> {
  const document = await lumaDb.documents.get(documentId);
  if (!document) {
    return {
      success: false,
      summary: "Project was not duplicated.",
      error: {
        code: "document_not_found",
        message: "LUMA could not find that Create project.",
        recoveryActions: ["Open Create"],
      },
    };
  }
  const copy: CreateDocument = {
    ...document,
    id: uid("doc"),
    ownerId: document.ownerId ?? ownerId(),
    title: `${document.title} Copy`,
    status: "draft",
  };
  const history = await recordAction({
    action: "duplicate_document",
    summary: `Duplicated "${document.title}".`,
    impact: "personal",
    sourceRefs: [asDocumentRef(document)],
    affected: asDocumentRef(copy),
    payload: { sourceDocumentId: document.id, copy },
    inversePayload: { deleteDocumentId: copy.id },
  });
  await lumaDb.documents.put(copy);
  return {
    success: true,
    data: copy,
    summary: `Duplicated "${document.title}".`,
    sourceRefs: [asDocumentRef(copy)],
    undo: {
      label: "Undo duplicate",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function deleteDocumentTool(documentId: string): Promise<AgentToolResult> {
  const document = await lumaDb.documents.get(documentId);
  if (!document) {
    return {
      success: false,
      summary: "Project was not deleted.",
      error: {
        code: "document_not_found",
        message: "LUMA could not find that Create project.",
        recoveryActions: ["Open Create"],
      },
    };
  }
  const history = await recordAction({
    action: "delete_document",
    summary: `Deleted project "${document.title}".`,
    impact: "destructive",
    sourceRefs: [asDocumentRef(document)],
    affected: asDocumentRef(document),
    payload: { documentId },
    inversePayload: { document },
  });
  await lumaDb.documents.delete(documentId);
  return {
    success: true,
    summary: `Deleted project "${document.title}".`,
    sourceRefs: [asDocumentRef(document)],
    undo: {
      label: "Undo delete",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Delete project",
      description: `This deletes "${document.title}" from Create.`,
      impact: "destructive",
    },
  };
}

export async function moveTaskTool(input: {
  taskId: string;
  naturalLanguageDate: string;
  now?: Date;
  provider?: AuditLogEntry["provider"];
}): Promise<AgentToolResult<TaskItem>> {
  const task = await lumaDb.tasks.get(input.taskId);
  if (!task) {
    return {
      success: false,
      summary: "Task was not moved.",
      error: {
        code: "task_not_found",
        message: "LUMA could not find that task in your local workspace.",
        recoveryActions: ["Open Calendar", "Search tasks"],
      },
    };
  }
  if (task.meta?.permissionState === "locked") {
    return {
      success: false,
      summary: "Task was not moved.",
      error: {
        code: "task_locked",
        message: "This task is locked and cannot be rescheduled from the agent.",
        recoveryActions: ["Open Task"],
      },
    };
  }
  const before = task.dueAt;
  const moved = {
    ...moveTaskByNaturalLanguage(task, input.naturalLanguageDate, input.now),
    meta: {
      ...(task.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local" as const,
    },
  };
  const history = await recordAction({
    action: "move_task",
    summary: `Moved "${moved.title}".`,
    impact: task.groupId ? "shared" : "personal",
    sourceRefs: [asTaskRef(task)],
    affected: asTaskRef(moved),
    payload: { taskId: task.id, dueAt: moved.dueAt },
    inversePayload: { taskId: task.id, dueAt: before },
    provider: input.provider ?? "manual",
  });
  await lumaDb.tasks.put(moved);
  return {
    success: true,
    data: moved,
    summary: `Moved "${moved.title}".`,
    sourceRefs: [asTaskRef(moved)],
    undo: {
      label: "Undo move",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateTaskTool(input: {
  taskId: string;
  patch: Partial<Pick<TaskItem, "title" | "subjectId" | "status" | "dueAt" | "priority" | "type" | "notes" | "estimatedMinutes" | "assignee">>;
  summary?: string;
}): Promise<AgentToolResult<TaskItem>> {
  const task = await lumaDb.tasks.get(input.taskId);
  if (!task) {
    return {
      success: false,
      summary: "Task was not updated.",
      error: {
        code: "task_not_found",
        message: "LUMA could not find that task in your local workspace.",
        recoveryActions: ["Open Calendar"],
      },
    };
  }
  const updated: TaskItem = {
    ...task,
    ...input.patch,
    meta: {
      ...(task.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local",
    },
  };
  const history = await recordAction({
    action: "update_task",
    summary: input.summary ?? `Updated task "${updated.title}".`,
    impact: task.groupId ? "shared" : "personal",
    sourceRefs: [asTaskRef(task)],
    affected: asTaskRef(updated),
    payload: { taskId: task.id, patch: input.patch },
    inversePayload: { task },
  });
  await lumaDb.tasks.put(updated);
  return {
    success: true,
    data: updated,
    summary: input.summary ?? `Updated task "${updated.title}".`,
    sourceRefs: [asTaskRef(updated)],
    undo: {
      label: "Undo task update",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function deleteTaskTool(taskId: string): Promise<AgentToolResult> {
  const task = await lumaDb.tasks.get(taskId);
  if (!task) {
    return {
      success: false,
      summary: "Task was not deleted.",
      error: {
        code: "task_not_found",
        message: "LUMA could not find that task in your local workspace.",
        recoveryActions: ["Open Calendar"],
      },
    };
  }
  const checklistItems = await lumaDb.checklistItems.where("taskId").equals(task.id).toArray();
  const history = await recordAction({
    action: "delete_task",
    summary: `Deleted task "${task.title}".`,
    impact: "destructive",
    sourceRefs: [asTaskRef(task)],
    affected: asTaskRef(task),
    payload: { taskId: task.id },
    inversePayload: { task, checklistItems },
  });
  await lumaDb.transaction("rw", lumaDb.tasks, lumaDb.checklistItems, async () => {
    await lumaDb.tasks.delete(task.id);
    await Promise.all(checklistItems.map((item) => lumaDb.checklistItems.delete(item.id)));
  });
  return {
    success: true,
    summary: `Deleted task "${task.title}".`,
    sourceRefs: [asTaskRef(task)],
    undo: {
      label: "Undo delete",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
    requiresConfirmation: true,
    confirmation: {
      title: "Delete task",
      description: `This deletes "${task.title}" and its checklist.`,
      impact: "destructive",
    },
  };
}

export async function createChecklistItemTool(input: {
  taskId: string;
  title: string;
}): Promise<AgentToolResult<ChecklistItem>> {
  const task = await lumaDb.tasks.get(input.taskId);
  const title = input.title.trim();
  if (!task || !title) {
    return {
      success: false,
      summary: "Checklist item was not created.",
      error: {
        code: !task ? "task_not_found" : "empty_checklist_title",
        message: !task ? "LUMA could not find that task." : "A checklist item needs text.",
        recoveryActions: ["Open task"],
      },
    };
  }
  const now = new Date().toISOString();
  const item: ChecklistItem = {
    id: uid("check"),
    ownerId: ownerId(),
    taskId: task.id,
    title,
    done: false,
    createdAt: now,
    updatedAt: now,
    syncState: "local",
    permissionState: "owner",
  };
  const updatedTask: TaskItem = {
    ...task,
    checklistItemIds: [...new Set([...(task.checklistItemIds ?? []), item.id])],
    meta: { ...(task.meta ?? createEntityMeta()), updatedAt: now, syncState: "local" },
  };
  const history = await recordAction({
    action: "create_checklist_item",
    summary: `Added checklist item to "${task.title}".`,
    impact: task.groupId ? "shared" : "personal",
    sourceRefs: [asTaskRef(task)],
    affected: asTaskRef(task),
    payload: { item },
    inversePayload: { deleteChecklistItemId: item.id, task },
  });
  await lumaDb.transaction("rw", lumaDb.tasks, lumaDb.checklistItems, async () => {
    await lumaDb.checklistItems.put(item);
    await lumaDb.tasks.put(updatedTask);
  });
  return {
    success: true,
    data: item,
    summary: `Added "${item.title}" to "${task.title}".`,
    sourceRefs: [asTaskRef(task)],
    undo: {
      label: "Undo checklist item",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateChecklistItemTool(input: {
  itemId: string;
  patch: Partial<Pick<ChecklistItem, "title" | "done">>;
}): Promise<AgentToolResult<ChecklistItem>> {
  const item = await lumaDb.checklistItems.get(input.itemId);
  if (!item) {
    return {
      success: false,
      summary: "Checklist item was not updated.",
      error: {
        code: "checklist_item_not_found",
        message: "LUMA could not find that checklist item.",
        recoveryActions: ["Open task"],
      },
    };
  }
  const updated: ChecklistItem = {
    ...item,
    ...input.patch,
    updatedAt: new Date().toISOString(),
    syncState: "local",
  };
  const history = await recordAction({
    action: "update_checklist_item",
    summary: `Updated checklist item "${updated.title}".`,
    impact: "personal",
    payload: { itemId: item.id, patch: input.patch },
    inversePayload: { checklistItem: item },
  });
  await lumaDb.checklistItems.put(updated);
  return {
    success: true,
    data: updated,
    summary: `Updated checklist item "${updated.title}".`,
    undo: {
      label: "Undo checklist update",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function updateTaskStatusTool(input: {
  taskId: string;
  status: TaskStatus;
}): Promise<AgentToolResult<TaskItem>> {
  const task = await lumaDb.tasks.get(input.taskId);
  if (!task) {
    return {
      success: false,
      summary: "Task was not updated.",
      error: {
        code: "task_not_found",
        message: "LUMA could not find that task in your local workspace.",
        recoveryActions: ["Open Calendar"],
      },
    };
  }
  const updated = {
    ...task,
    status: input.status,
    meta: {
      ...(task.meta ?? createEntityMeta()),
      updatedAt: new Date().toISOString(),
      syncState: "local" as const,
    },
  };
  const history = await recordAction({
    action: "update_task_status",
    summary: `Updated "${updated.title}" to ${updated.status}.`,
    impact: task.groupId ? "shared" : "personal",
    affected: asTaskRef(updated),
    payload: { taskId: task.id, status: updated.status },
    inversePayload: { taskId: task.id, status: task.status },
  });
  await lumaDb.tasks.put(updated);
  return {
    success: true,
    data: updated,
    summary: `Updated "${updated.title}" to ${updated.status}.`,
    sourceRefs: [asTaskRef(updated)],
    undo: {
      label: "Undo status change",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function startFocusSessionTool(input: {
  subjectId: string;
  minutes: number;
  mode?: StudySession["mode"];
  taskId?: string;
  provider?: AuditLogEntry["provider"];
}): Promise<AgentToolResult<StudySession>> {
  const now = new Date().toISOString();
  const session: StudySession = {
    id: uid("focus"),
    subjectId: input.subjectId,
    minutes: input.minutes,
    mode: input.mode ?? "deep-work",
    startedAt: now,
    completed: false,
    taskId: input.taskId,
    meta: createEntityMeta(ownerId(), now),
  };
  const sourceRefs: SourceReference[] = [
    {
      type: "focus",
      id: session.id,
      title: `${session.minutes} minute focus session`,
      location: session.startedAt,
    },
  ];
  const history = await recordAction({
    action: "start_focus_session",
    summary: `Started a ${session.minutes} minute focus session.`,
    impact: "personal",
    sourceRefs,
    affected: sourceRefs[0],
    payload: { session },
    inversePayload: { deleteStudySessionId: session.id },
    provider: input.provider ?? "manual",
  });
  await lumaDb.studySessions.put(session);
  return {
    success: true,
    data: session,
    summary: `Started a ${session.minutes} minute focus session.`,
    sourceRefs,
    undo: {
      label: "Undo focus start",
      actionId: history.id,
      expiresAt: history.expiresAt,
    },
  };
}

export async function undoAction(actionId: string): Promise<AgentToolResult> {
  const action = await lumaDb.actionHistory.get(actionId);
  if (!action || action.status !== "completed" || !action.inversePayload) {
    return {
      success: false,
      summary: "Nothing was undone.",
      error: {
        code: "undo_unavailable",
        message: "That action cannot be undone anymore.",
        recoveryActions: ["Open action history"],
      },
    };
  }

  if (typeof action.inversePayload.deleteTaskId === "string") {
    await lumaDb.tasks.delete(action.inversePayload.deleteTaskId);
  }
  if (typeof action.inversePayload.deleteChecklistItemId === "string") {
    await lumaDb.checklistItems.delete(action.inversePayload.deleteChecklistItemId);
  }
  if (typeof action.inversePayload.deleteStudySessionId === "string") {
    await lumaDb.studySessions.delete(action.inversePayload.deleteStudySessionId);
  }
  if (typeof action.inversePayload.deleteGroupId === "string") {
    await lumaDb.groups.delete(action.inversePayload.deleteGroupId);
  }
  if (typeof action.inversePayload.deleteDocumentId === "string") {
    await lumaDb.documents.delete(action.inversePayload.deleteDocumentId);
  }
  if (typeof action.inversePayload.deleteDocumentVersionId === "string") {
    await lumaDb.documentVersions.delete(action.inversePayload.deleteDocumentVersionId);
  }
  if (typeof action.inversePayload.deleteMaterialId === "string") {
    await lumaDb.materials.delete(action.inversePayload.deleteMaterialId);
  }
  if (typeof action.inversePayload.deleteHighlightId === "string") {
    await lumaDb.highlights.delete(action.inversePayload.deleteHighlightId);
  }
  if (Array.isArray(action.inversePayload.deleteFlashcardIds)) {
    await Promise.all(
      action.inversePayload.deleteFlashcardIds
        .filter((item): item is string => typeof item === "string")
        .map((id) => lumaDb.flashcards.delete(id)),
    );
  }
  if (typeof action.inversePayload.deleteFlashcardReviewId === "string") {
    await lumaDb.flashcardReviews.delete(action.inversePayload.deleteFlashcardReviewId);
  }
  if (typeof action.inversePayload.deleteQuizId === "string") {
    await lumaDb.quizzes.delete(action.inversePayload.deleteQuizId);
  }
  if (typeof action.inversePayload.deleteQuizAttemptId === "string") {
    await lumaDb.quizAttempts.delete(action.inversePayload.deleteQuizAttemptId);
  }
  if (typeof action.inversePayload.deleteChatId === "string") {
    await lumaDb.chats.delete(action.inversePayload.deleteChatId);
  }
  if (typeof action.inversePayload.materialId === "string" && Array.isArray(action.inversePayload.sharedWith)) {
    const material = await lumaDb.materials.get(action.inversePayload.materialId);
    if (material) {
      await lumaDb.materials.put({
        ...material,
        sharedWith: action.inversePayload.sharedWith.filter((item): item is string => typeof item === "string"),
        updatedAt: new Date().toISOString(),
      });
    }
  }
  if ("document" in action.inversePayload && action.inversePayload.document && typeof action.inversePayload.document === "object") {
    await lumaDb.documents.put(action.inversePayload.document as CreateDocument);
  }
  if ("group" in action.inversePayload && action.inversePayload.group && typeof action.inversePayload.group === "object") {
    await lumaDb.groups.put(action.inversePayload.group as GroupWorkspace);
  }
  if ("chat" in action.inversePayload && action.inversePayload.chat && typeof action.inversePayload.chat === "object") {
    await lumaDb.chats.put(action.inversePayload.chat as ChatMessage);
  }
  if ("material" in action.inversePayload && action.inversePayload.material && typeof action.inversePayload.material === "object") {
    await lumaDb.materials.put(action.inversePayload.material as Material);
  }
  if ("flashcard" in action.inversePayload && action.inversePayload.flashcard && typeof action.inversePayload.flashcard === "object") {
    await lumaDb.flashcards.put(action.inversePayload.flashcard as Flashcard);
  }
  if ("quiz" in action.inversePayload && action.inversePayload.quiz && typeof action.inversePayload.quiz === "object") {
    await lumaDb.quizzes.put(action.inversePayload.quiz as Quiz);
  }
  if (Array.isArray(action.inversePayload.highlights)) {
    await lumaDb.highlights.bulkPut(action.inversePayload.highlights as Highlight[]);
  }
  if ("task" in action.inversePayload && action.inversePayload.task && typeof action.inversePayload.task === "object") {
    await lumaDb.tasks.put(action.inversePayload.task as TaskItem);
  }
  if (Array.isArray(action.inversePayload.checklistItems)) {
    await lumaDb.checklistItems.bulkPut(action.inversePayload.checklistItems as ChecklistItem[]);
  }
  if ("checklistItem" in action.inversePayload && action.inversePayload.checklistItem && typeof action.inversePayload.checklistItem === "object") {
    await lumaDb.checklistItems.put(action.inversePayload.checklistItem as ChecklistItem);
  }
  if (typeof action.inversePayload.taskId === "string" && "dueAt" in action.inversePayload) {
    const task = await lumaDb.tasks.get(action.inversePayload.taskId);
    if (task && typeof action.inversePayload.dueAt === "string") {
      await lumaDb.tasks.put({
        ...task,
        dueAt: action.inversePayload.dueAt,
        meta: {
          ...(task.meta ?? createEntityMeta()),
          updatedAt: new Date().toISOString(),
          syncState: "local",
        },
      });
    }
  }
  if (typeof action.inversePayload.taskId === "string" && "status" in action.inversePayload) {
    const task = await lumaDb.tasks.get(action.inversePayload.taskId);
    if (task && typeof action.inversePayload.status === "string") {
      await lumaDb.tasks.put({
        ...task,
        status: action.inversePayload.status as TaskStatus,
        meta: {
          ...(task.meta ?? createEntityMeta()),
          updatedAt: new Date().toISOString(),
          syncState: "local",
        },
      });
    }
  }

  const now = new Date().toISOString();
  await lumaDb.actionHistory.put({ ...action, status: "undone", updatedAt: now });
  await lumaDb.auditLog.put({
    id: uid("audit"),
    ownerId: ownerId(),
    action: "undo_action",
    summary: `Undid: ${action.summary}`,
    impact: "personal",
    sourceRefs: [],
    provider: "manual",
    createdAt: now,
  });

  return {
    success: true,
    summary: `Undid: ${action.summary}`,
  };
}
