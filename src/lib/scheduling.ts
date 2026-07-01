import { addDays, differenceInMinutes, format, isBefore, parseISO, startOfToday } from "date-fns";
import type { ClassSession, Flashcard, Subject, TaskItem } from "./types";

function asDate(value: string | Date) {
  return value instanceof Date ? value : parseISO(value);
}

export function subjectById(subjects: Subject[], id: string) {
  return subjects.find((subject) => subject.id === id);
}

export function minutesUntil(dateIso: string, now = new Date()) {
  return Math.max(0, differenceInMinutes(parseISO(dateIso), now));
}

export function formatDue(dateIso: string | Date) {
  const due = asDate(dateIso);
  const today = startOfToday();
  if (format(due, "yyyy-MM-dd") === format(today, "yyyy-MM-dd")) {
    return `Today, ${format(due, "p")}`;
  }
  if (format(due, "yyyy-MM-dd") === format(addDays(today, 1), "yyyy-MM-dd")) {
    return `Tomorrow, ${format(due, "p")}`;
  }
  return format(due, "MMM d, p");
}

export function tasksDueWithin(tasks: TaskItem[], days: number, now = new Date()) {
  const limit = addDays(now, days);
  return tasks
    .filter((task) => task.status !== "done")
    .filter((task) => isBefore(asDate(task.dueAt), limit))
    .sort((a, b) => asDate(a.dueAt).getTime() - asDate(b.dueAt).getTime());
}

export function getNextClass(sessions: ClassSession[], now = new Date()) {
  const candidates = sessions.map((session) => {
    const [hours, minutes] = session.start.split(":").map(Number);
    const date = new Date(now);
    const diff = (session.weekday - now.getDay() + 7) % 7;
    date.setDate(now.getDate() + diff);
    date.setHours(hours, minutes, 0, 0);
    if (date.getTime() < now.getTime()) {
      date.setDate(date.getDate() + 7);
    }
    return { session, startsAt: date.toISOString(), minutesAway: differenceInMinutes(date, now) };
  });

  return candidates.sort((a, b) => a.minutesAway - b.minutesAway)[0];
}

export function moveTaskByNaturalLanguage(task: TaskItem, input: string, now = new Date()): TaskItem {
  const lower = input.toLowerCase();
  const due = parseISO(task.dueAt);
  if (lower.includes("tomorrow")) {
    const base = new Date(now);
    base.setHours(0, 0, 0, 0);
    const next = addDays(base, 1);
    next.setHours(due.getHours(), due.getMinutes(), 0, 0);
    return { ...task, dueAt: next.toISOString() };
  }
  if (lower.includes("friday")) {
    const target = new Date(now);
    const diff = (5 - now.getDay() + 7) % 7 || 7;
    target.setDate(now.getDate() + diff);
    target.setHours(due.getHours(), due.getMinutes(), 0, 0);
    return { ...task, dueAt: target.toISOString() };
  }
  return { ...task, dueAt: addDays(due, 1).toISOString() };
}

export function gradeFlashcard(card: Flashcard, quality: "again" | "hard" | "good" | "easy", now = new Date()) {
  const delta = quality === "again" ? -24 : quality === "hard" ? 24 : quality === "good" ? 72 : 168;
  const masteryDelta = quality === "again" ? -18 : quality === "hard" ? 6 : quality === "good" ? 14 : 24;
  return {
    ...card,
    mastery: Math.min(100, Math.max(0, card.mastery + masteryDelta)),
    dueAt: addDays(now, Math.max(1, Math.round(delta / 24))).toISOString(),
  };
}
