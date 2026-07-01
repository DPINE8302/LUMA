import { createSeedData } from "../data/seed";
import { DEFAULT_OWNER_ID, lumaDb, setCurrentOwnerId } from "./db";
import type { AuthSession, LocalAccount, OnboardingState, UserProfile, UserSettings } from "./types";

const encoder = new TextEncoder();

function uid(prefix: string) {
  return `${prefix}-${crypto.randomUUID()}`;
}

function normalizeEmail(email: string) {
  return email.trim().toLowerCase();
}

function initials(name: string) {
  return name
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0]?.toUpperCase())
    .join("") || "LU";
}

async function digest(value: string) {
  const hash = await crypto.subtle.digest("SHA-256", encoder.encode(value));
  return Array.from(new Uint8Array(hash))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

async function hashPassword(password: string, salt: string = crypto.randomUUID()) {
  return {
    salt,
    hash: await digest(`${salt}:${password}`),
  };
}

async function deactivateSessions() {
  const active = (await lumaDb.authSessions.toArray()).filter((session) => session.active);
  const now = new Date().toISOString();
  await lumaDb.authSessions.bulkPut(active.map((session) => ({ ...session, active: false, updatedAt: now })));
}

export async function signInLocalAccount(email: string, password: string): Promise<{ success: true; session: AuthSession } | { success: false; message: string }> {
  const account = await lumaDb.accounts.where("email").equals(normalizeEmail(email)).first();
  if (!account) {
    return { success: false, message: "No local LUMA account uses that email." };
  }
  const passwordHash = await hashPassword(password, account.passwordSalt);
  if (passwordHash.hash !== account.passwordHash) {
    return { success: false, message: "The password does not match this local account." };
  }
  const now = new Date().toISOString();
  await deactivateSessions();
  const session: AuthSession = {
    id: uid("session"),
    accountId: account.id,
    profileId: account.profileId,
    active: true,
    createdAt: now,
    updatedAt: now,
  };
  await lumaDb.transaction("rw", lumaDb.accounts, lumaDb.authSessions, async () => {
    await lumaDb.accounts.put({ ...account, lastSignedInAt: now, updatedAt: now });
    await lumaDb.authSessions.put(session);
  });
  setCurrentOwnerId(account.profileId);
  return { success: true, session };
}

export async function createLocalAccount(input: {
  name: string;
  email: string;
  password: string;
  school?: string;
  gradeOrYear?: string;
  program?: string;
}): Promise<{ success: true; session: AuthSession } | { success: false; message: string }> {
  const email = normalizeEmail(input.email);
  if (!email.includes("@")) return { success: false, message: "Use a valid email for this local account." };
  if (input.password.length < 8) return { success: false, message: "Use at least 8 characters for the local password." };
  const existing = await lumaDb.accounts.where("email").equals(email).first();
  if (existing) return { success: false, message: "A local LUMA account already exists for that email." };

  const seed = createSeedData();
  const now = new Date().toISOString();
  const profileId = uid("profile");
  const accountId = uid("account");
  const passwordHash = await hashPassword(input.password);
  const profile: UserProfile = {
    ...seed.profile,
    id: profileId,
    name: input.name.trim() || "LUMA Student",
    username: email.split("@")[0],
    avatar: initials(input.name),
    school: input.school?.trim() || "",
    gradeOrYear: input.gradeOrYear?.trim() || "",
    program: input.program?.trim() || "",
    year: [input.gradeOrYear, input.program].filter(Boolean).join(" · ") || "Student",
    focusSubjects: [],
    goals: [],
    achievements: [],
    journey: [],
    preferencesSnapshot: [],
    currentGoal: undefined,
    metricsVisible: seed.profile.metricsVisible,
  };
  const account: LocalAccount = {
    id: accountId,
    profileId,
    email,
    displayName: profile.name,
    passwordHash: passwordHash.hash,
    passwordSalt: passwordHash.salt,
    mode: "local",
    createdAt: now,
    updatedAt: now,
    lastSignedInAt: now,
    syncState: "local",
    permissionState: "owner",
  };
  const session: AuthSession = {
    id: uid("session"),
    accountId,
    profileId,
    active: true,
    createdAt: now,
    updatedAt: now,
  };
  const onboarding: OnboardingState = {
    id: uid("onboarding"),
    accountId,
    profileId,
    completed: false,
    skipped: false,
    currentStep: 1,
    completedSteps: ["profile"],
    createdAt: now,
    updatedAt: now,
  };
  const settings: UserSettings = {
    ...seed.settings,
    id: `settings-${profileId}`,
  };

  await deactivateSessions();
  await lumaDb.transaction("rw", lumaDb.accounts, lumaDb.authSessions, lumaDb.onboarding, lumaDb.profile, lumaDb.settings, async () => {
    await lumaDb.profile.put(profile);
    await lumaDb.settings.put(settings);
    await lumaDb.accounts.put(account);
    await lumaDb.authSessions.put(session);
    await lumaDb.onboarding.put(onboarding);
  });
  setCurrentOwnerId(profileId);
  return { success: true, session };
}

export async function signOutLocalAccount() {
  await deactivateSessions();
  setCurrentOwnerId(DEFAULT_OWNER_ID);
}

export async function resetLocalPassword(email: string, nextPassword: string): Promise<{ success: boolean; message: string }> {
  if (nextPassword.length < 8) return { success: false, message: "Use at least 8 characters for the new password." };
  const account = await lumaDb.accounts.where("email").equals(normalizeEmail(email)).first();
  if (!account) return { success: false, message: "No local LUMA account uses that email." };
  if (account.mode === "demo") return { success: false, message: "The demo account does not use password reset." };
  const passwordHash = await hashPassword(nextPassword);
  await lumaDb.accounts.put({ ...account, passwordHash: passwordHash.hash, passwordSalt: passwordHash.salt, updatedAt: new Date().toISOString() });
  return { success: true, message: "Password updated for this local device." };
}

export async function completeOnboarding(accountId: string, stepIds: string[]) {
  const onboarding = await lumaDb.onboarding.where("accountId").equals(accountId).first();
  if (!onboarding) return;
  await lumaDb.onboarding.put({
    ...onboarding,
    completed: true,
    skipped: false,
    currentStep: 10,
    completedSteps: Array.from(new Set([...onboarding.completedSteps, ...stepIds])),
    updatedAt: new Date().toISOString(),
  });
}

export async function skipOnboarding(accountId: string) {
  const onboarding = await lumaDb.onboarding.where("accountId").equals(accountId).first();
  if (!onboarding) return;
  await lumaDb.onboarding.put({
    ...onboarding,
    completed: false,
    skipped: true,
    currentStep: Math.max(onboarding.currentStep, 2),
    updatedAt: new Date().toISOString(),
  });
}
