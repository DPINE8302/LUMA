import { expect, test, type Page } from "@playwright/test";

async function openRoute(page: Page, route: string) {
  const navToggle = page.getByRole("button", { name: /Open navigation|Close navigation/ });
  const viewport = page.viewportSize();
  const useMobileNav = (viewport?.width ?? 1280) <= 900 || (await navToggle.isVisible({ timeout: 800 }).catch(() => false));
  if (route === "Settings") {
    const settings = page.getByRole("button", { name: "Settings" }).first();
    if (useMobileNav) {
      await navToggle.click({ force: true });
      await expect(page.locator(".responsive-rail")).toBeVisible();
    }
    await expect(settings).toBeVisible();
    await settings.click();
    return;
  }
  const desktopTarget = page.locator(`.side-rail button[data-route="${route}"]`).first();
  if (!useMobileNav) {
    await expect(desktopTarget).toBeVisible();
    await desktopTarget.click();
    return;
  }
  await navToggle.click({ force: true });
  await expect(page.locator(".responsive-rail")).toBeVisible();
  const mobileTarget = page.locator(`.responsive-rail button[data-route="${route}"]`).first();
  await expect(mobileTarget).toBeVisible();
  await mobileTarget.click();
}

async function openTogetherTab(page: Page, tab: string) {
  const mobileTab = page.locator(".together-mobile-tabs").getByRole("button", { name: tab });
  if (await mobileTab.isVisible().catch(() => false)) {
    await mobileTab.click();
  }
}

async function expectNoHorizontalOverflow(page: Page) {
  const overflow = await page.evaluate(() => {
    const root = document.documentElement;
    return Math.ceil(root.scrollWidth - root.clientWidth);
  });
  expect(overflow).toBeLessThanOrEqual(2);
}

async function expectCoreChromeVisible(page: Page, options: { agentVisible?: boolean } = {}) {
  await expect
    .poll(async () => {
      const hasHamburger = await page.getByRole("button", { name: /Open navigation|Close navigation/ }).isVisible().catch(() => false);
      const hasRail = await page.locator(".side-rail").isVisible().catch(() => false);
      return hasHamburger || hasRail;
    })
    .toBe(true);
  if (options.agentVisible ?? true) {
    await expect(page.getByPlaceholder(/Ask LUMA anything/)).toBeVisible();
  }
  await expectNoHorizontalOverflow(page);
}

test("renders LUMA home and navigates main areas", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Good morning, Win" })).toBeVisible();
  await expect(page.getByText("Today's Brief").first()).toBeVisible();
  await expect(page.getByText(/Next Focus|No normal subject classes|No normal classes/).first()).toBeVisible();
  await expect(page.getByText(/3D Model Sculpting|Art 9|Game Design & Development|Additional Mathematics 3|No normal classes/).first()).toBeVisible();

  await openRoute(page, "Learn");
  await expect(page.getByRole("heading", { name: "Learn", exact: true })).toBeVisible();
  await expect(page.getByText("Continue Learning")).toBeVisible();
  await expect(page.getByRole("heading", { name: "Data Structures Notes - Search Algorithms" })).toBeVisible();

  await openRoute(page, "Calendar");
  await expect(page.getByRole("heading", { name: /Calendar/ })).toBeVisible();
  await expect(page.getByText("Smart Timetable")).toBeVisible();
  await expect(page.getByText("Art 9").first()).toBeVisible();
  await expect(page.getByText("Game Design & Development").first()).toBeVisible();
  await expect(page.getByText("Additional Mathematics 3").first()).toBeVisible();
  await expect(page.getByText("No normal classes").first()).toBeVisible();

  await openRoute(page, "Together");
  await expect(page.getByRole("heading", { name: "Together" })).toBeVisible();
  await expect(page.getByRole("heading", { name: /IM Capstone Crew/ })).toBeVisible();

  await openRoute(page, "Create");
  await expect(page.getByRole("heading", { name: "Create", exact: true })).toBeVisible();
  await expect(page.getByText("Document Creation Flow")).toBeVisible();
  await page.getByRole("button", { name: /New project/ }).click();
  await expect(page.getByText(/Created project/)).toBeVisible();
  await expect(page.getByText("Untitled Project").first()).toBeVisible();

  await openRoute(page, "Profile");
  await expect(page.getByRole("heading", { name: "Profile", exact: true })).toBeVisible();
  await expect(page.getByText("Study Analytics")).toBeVisible();
  await expect(page.getByText("No achievements yet.")).toBeVisible();
});

test("creates a local account and completes first-use onboarding", async ({ page }) => {
  await page.goto("/");
  await openRoute(page, "Settings");
  await page.getByRole("button", { name: /Backup & Sync/ }).click();
  page.once("dialog", async (dialog) => {
    await dialog.accept();
  });
  await page.getByRole("button", { name: /Sign Out/ }).click();
  await expect(page.getByRole("heading", { name: "Welcome back" })).toBeVisible();

  await page.getByRole("button", { name: "Create account" }).click();
  await page.getByPlaceholder("Your name").fill("Maya Chen");
  await page.getByPlaceholder("you@school.edu").fill(`maya-${Date.now()}@school.test`);
  await page.getByPlaceholder("At least 8 characters").fill("quietpass1");
  await page.getByRole("textbox", { name: "School" }).fill("LUMA Academy");
  await page.getByRole("textbox", { name: "Grade" }).fill("M.5");
  await page.getByRole("textbox", { name: "Program" }).fill("Science");
  await page.getByRole("button", { name: "Create Account", exact: true }).click();

  await expect(page.getByRole("heading", { name: "Make this space yours" })).toBeVisible();
  await page.getByRole("button", { name: "Finish Setup" }).click();
  await expect(page.getByRole("heading", { name: "Good morning, Maya" })).toBeVisible();

  await page.reload();
  await expect(page.getByRole("heading", { name: "Good morning, Maya" })).toBeVisible();
  await openRoute(page, "Profile");
  await expect(page.getByRole("heading", { name: "Maya Chen" })).toBeVisible();
});

test("creates a task, persists after reload, and uses the global LUMA agent fallback", async ({ page }) => {
  await page.goto("/");
  await openRoute(page, "Calendar");
  await page.getByRole("button", { name: "Add item" }).click();
  await page.getByPlaceholder("Add homework, exam, or project...").fill("Finish Art 9 online assignment");
  await page.getByRole("button", { name: /^Add$/ }).click();
  await expect(page.getByText("Finish Art 9 online assignment")).toBeVisible();
  await page.getByRole("button", { name: /Finish Art 9 online assignment/ }).click();
  await expect(page.getByRole("heading", { name: "Task Detail" })).toBeVisible();
  await page.getByLabel("Title").fill("Finish Art 9 online assignment revised");
  await page.getByPlaceholder("Add a step...").fill("Upload worksheet screenshot");
  await page.getByPlaceholder("Add a step...").press("Enter");
  await expect(page.getByText("Upload worksheet screenshot", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "Save Task" }).click();
  await expect(page.getByText(/Updated task/)).toBeVisible();
  await page.getByRole("button", { name: "Complete" }).click();
  await expect(page.getByText(/Updated/)).toBeVisible();
  await page.getByRole("button", { name: "Reopen" }).click();
  await expect(page.getByText(/Updated/)).toBeVisible();
  await page.getByRole("button", { name: "Close task detail" }).click();

  await page.reload();
  await openRoute(page, "Calendar");
  await expect(page.getByText("Finish Art 9 online assignment revised")).toBeVisible();

  await page.getByPlaceholder(/Ask LUMA anything/).fill("What is due this week? Make flashcards.");
  await page.locator(".global-agent button[type='submit']").click();
  await expect(page.locator(".agent-result-head").getByText("LUMA")).toBeVisible();
  await expect(page.locator(".agent-popover").getByText("Using")).toBeVisible();
  await expect(page.locator(".agent-context").getByRole("button", { name: "Calendar" })).toBeVisible();
  await expect(page.locator(".agent-message-sources button").first()).toBeVisible();
  await expect(page.getByText(/active item|No urgent work|Upload or create a study note first|I can/)).toBeVisible();
});

test("creates a task through the LUMA Agent plan preview", async ({ page }) => {
  await page.goto("/");
  await page.getByPlaceholder(/Ask LUMA anything/).fill("Add task finish Art 9 reflection tomorrow");
  await page.locator(".global-agent button[type='submit']").click();
  await expect(page.locator(".agent-actions").getByRole("button", { name: /Create task: finish Art 9 reflection/ })).toBeVisible();
  await page.locator(".agent-actions").getByRole("button", { name: /Create task: finish Art 9 reflection/ }).click();
  await expect(page.locator(".agent-confirm").getByText("Plan preview")).toBeVisible();
  await expect(page.locator(".agent-confirm").getByText(/Create a private task/)).toBeVisible();
  await page.locator(".agent-confirm").getByRole("button", { name: "Run action" }).click();
  await expect(page.locator(".agent-popover").getByText(/Created task "finish Art 9 reflection"/)).toBeVisible();
  await expect(page.locator(".agent-message-sources").getByRole("button", { name: "finish Art 9 reflection" })).toBeVisible();
  await page.locator(".agent-message-sources").getByRole("button", { name: "finish Art 9 reflection" }).click();
  await expect(page.getByRole("heading", { name: /Calendar/ })).toBeVisible();
  await expect(page.getByRole("button", { name: /finish Art 9 reflection Tomorrow/ })).toBeVisible();

  await page.reload();
  await openRoute(page, "Calendar");
  await expect(page.getByRole("button", { name: /finish Art 9 reflection Tomorrow/ })).toBeVisible();
});

test("uploads and manages Study Vault material from Learn", async ({ page }) => {
  await page.goto("/");
  await openRoute(page, "Learn");
  await page.locator('input[type="file"]').setInputFiles({
    name: "algorithms.txt",
    mimeType: "text/plain",
    buffer: Buffer.from("Binary search runs in logarithmic time. Use it on sorted arrays."),
  });
  await expect(page.getByText(/Added "algorithms" to Study Vault/)).toBeVisible();
  await expect(page.getByRole("heading", { name: "algorithms" })).toBeVisible();

  await page.getByLabel("Material folder").fill("Computer Science");
  await page.getByRole("button", { name: "Move" }).click();
  await expect(page.getByText(/Moved "algorithms" to Computer Science/)).toBeVisible();

  await page.getByLabel("Material tag").fill("search");
  await page.getByRole("button", { name: "Tag" }).click();
  await expect(page.getByText(/Tagged "algorithms" with search/)).toBeVisible();

  await page.getByRole("button", { name: /Save Highlight/ }).click();
  await expect(page.getByText(/Saved a highlight/)).toBeVisible();

  await page.getByRole("button", { name: /Create Cards/ }).click();
  await expect(page.getByText(/Created 2 flashcards/)).toBeVisible();
  await page.getByRole("button", { name: "Reveal answer" }).click();
  await expect(page.getByRole("button", { name: "Hide answer" })).toBeVisible();
  await page.getByRole("button", { name: "easy" }).click();
  await expect(page.getByText(/Reviewed/)).toBeVisible();

  await page.getByRole("button", { name: "Generate Quiz" }).click();
  await expect(page.getByText(/Created quiz/)).toBeVisible();
  await page.getByRole("combobox").selectOption("The key concept from the selected material");
  await page.getByPlaceholder("Write your answer").fill("It helps connect the material to the exam goal.");
  await page.getByRole("button", { name: "Submit Quiz" }).click();
  await expect(page.getByText(/Submitted/)).toBeVisible();
  await expect(page.locator(".quiz-score-ring").getByText("100%", { exact: true })).toBeVisible();

  await page.reload();
  await openRoute(page, "Learn");
  await expect(page.getByText("algorithms").first()).toBeVisible();
  await expect(page.getByText("2 cards")).toBeVisible();
  await expect(page.locator(".quiz-score-ring").getByText("100%", { exact: true })).toBeVisible();
});

test("creates, edits, versions, and restores Create projects", async ({ page }) => {
  await page.goto("/");
  await openRoute(page, "Create");
  await page.getByRole("button", { name: /New project/ }).click();
  await expect(page.getByText(/Created project/)).toBeVisible();

  const editor = page.locator(".document-page textarea");
  await editor.fill("First saved draft for the version history test.");
  await expect(page.getByRole("heading", { name: "Version History" })).toBeVisible();
  await expect(page.locator(".version-panel button").first()).toBeVisible();

  await page.getByLabel("Outline section 1").fill("Updated introduction");
  await expect(page.getByText(/Updated outline/)).toBeVisible();

  await page.reload();
  await openRoute(page, "Create");
  await expect(page.locator(".document-page textarea")).toHaveValue(/First saved draft/);
  await expect(page.getByLabel("Outline section 1")).toHaveValue("Updated introduction");
  await expect(page.locator(".version-panel button").first()).toBeVisible();
});

test("manages Together collaboration actions", async ({ page }) => {
  await page.goto("/");
  await openRoute(page, "Learn");
  await page.locator('input[type="file"]').setInputFiles({
    name: "team-notes.txt",
    mimeType: "text/plain",
    buffer: Buffer.from("Shared project notes for group work."),
  });
  await expect(page.getByText(/Added "team-notes" to Study Vault/)).toBeVisible();

  await openRoute(page, "Together");
  await expect(page.getByRole("heading", { name: /IM Capstone Crew/ })).toBeVisible();

  await openTogetherTab(page, "Files");
  page.once("dialog", async (dialog) => {
    await dialog.accept();
  });
  await page.getByRole("button", { name: /Share latest/ }).click();
  await expect(page.getByText(/Shared "team-notes"/)).toBeVisible();
  await expect(page.getByRole("button", { name: "Remove" }).first()).toBeVisible();

  await openTogetherTab(page, "Chat");
  await page.getByPlaceholder("Message group...").fill("Initial group update");
  await page.getByTitle("Send message").click();
  await expect(page.getByText("Initial group update")).toBeVisible();

  page.once("dialog", async (dialog) => {
    await dialog.accept("Edited group update");
  });
  await page.locator(".chat-panel article").filter({ hasText: "Initial group update" }).getByRole("button", { name: "Edit" }).click();
  await expect(page.getByText("Edited group update")).toBeVisible();

  await openTogetherTab(page, "Tasks");
  const taskBoard = page.locator(".task-board-panel");
  await taskBoard.getByRole("button", { name: /Add task/ }).first().click();
  await expect(page.locator(".task-board-panel").getByText("Summarize project progress")).toBeVisible();
  const groupTaskCard = taskBoard.locator("article").filter({ hasText: "Summarize project progress" }).first();
  await groupTaskCard.getByRole("button", { name: "Start" }).click();
  await expect(page.getByText(/Updated "Summarize project progress"/)).toBeVisible();
  await expect(taskBoard.locator(".progress").getByText("Summarize project progress")).toBeVisible();

  await openTogetherTab(page, "Overview");
  page.once("dialog", async (dialog) => {
    await dialog.accept();
  });
  await page.locator(".role-select-row select").selectOption("admin");
  await expect(page.getByText(/Changed You's role to admin/)).toBeVisible();
});

test("mobile home keeps the command space usable", async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Good morning, Win" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Open navigation" })).toBeVisible();
  await expect(page.getByPlaceholder(/Ask LUMA anything/)).toBeVisible();
  await expectNoHorizontalOverflow(page);
});

test("core routes keep primary chrome usable across Safari-sized viewports", async ({ page }, testInfo) => {
  if (testInfo.project.name === "chromium") {
    await page.setViewportSize({ width: 1440, height: 1200 });
  }

  await page.goto("/");
  await expectCoreChromeVisible(page);

  for (const route of ["Calendar", "Learn", "Together", "Create", "Profile", "Settings"]) {
    await openRoute(page, route);
    const usesCompactAgentDock = ["Learn", "Create"].includes(route);
    await expectCoreChromeVisible(page, { agentVisible: !usesCompactAgentDock });
    if (usesCompactAgentDock) {
      await expect(page.locator(".global-agent.is-compact-dock")).toBeVisible();
      await page.keyboard.press("Control+K");
      await expect(page.getByPlaceholder(/Ask LUMA anything/)).toBeVisible();
      await expect(page.locator(".agent-popover")).toBeVisible();
      await page.getByRole("button", { name: "Close LUMA" }).click();
    }
  }
});
