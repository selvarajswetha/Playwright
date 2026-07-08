const { test } = require("@playwright/test");

const fs = require("fs");

test("Login and Save Auth State", async ({ browser }) => {
  const context = await browser.newContext();

  const page = await context.newPage();

  await page.goto(
    "https://luvcarwashnew.sharepoint.com/sites/luvdowntimetracking/SitePages/OfflineAssetWorkOrders.aspx",
  );

  // Login

  await page
    .getByPlaceholder("Email, phone, or Skype")
    .fill("tracking@luvcarwash.com");

  await page.getByRole("button", { name: "Next" }).click();

  await page.getByPlaceholder("Password").fill("Q!743142452125up");

  await page.getByRole("button", { name: "Sign in" }).click();

  await page.getByRole("button", { name: "Yes" }).click();

  // Save authentication

  await context.storageState({ path: "tests/authState.json" });

  console.log("Auth state saved.");

  await browser.close();
});
