const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");
const { google } = require("googleapis");
require("dotenv").config();

const DEBUG = !!process.env.DEBUG;
function dlog(...args) {
  if (DEBUG) console.log("[DEBUG]", ...args);
}

const CHANNEL_URL =
  "https://discord.com/channels/1382818633348026501/1382818633348026504";
const SHEET_ID = "1ibOWAjmOZMeEBtetqKq6-0wmxDM0eHWj8EsLSAu0Xc0";
const TAB_NAME = "RECIPES";
const PROMPT_COLUMN = "K";
const START_ROW = 2;
if (!SHEET_ID) {
  console.error(
    "Please set SHEET_ID environment variable to your Google Sheet ID."
  );
  process.exit(1);
}

// Sleeps for 2 secs
function sleepTwoSecs() {
  return new Promise((resolve) => setTimeout(resolve, 2 * 1000));
}

// Array for the links of the images
const imageLinksArray = [];

/* ========== Google Sheets helpers ========== */
async function sheetsClient() {
  if (
    !process.env.GOOGLE_APPLICATION_CREDENTIALS ||
    !fs.existsSync(process.env.GOOGLE_APPLICATION_CREDENTIALS)
  ) {
    throw new Error(
      "Set GOOGLE_APPLICATION_CREDENTIALS env var to your service account JSON key file path and make sure the file exists."
    );
  }
  const auth = new google.auth.GoogleAuth({
    keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  const authClient = await auth.getClient();
  return google.sheets({ version: "v4", auth: authClient });
}

async function readSheetPrompts(sheets) {
  const rangeK = `${TAB_NAME}!${PROMPT_COLUMN}${START_ROW}:${PROMPT_COLUMN}`;
  const rangeD = `${TAB_NAME}!D${START_ROW}:D`;
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID,
    ranges: [rangeK, rangeD],
    majorDimension: "COLUMNS",
  });

  const promptValues =
    res.data.valueRanges &&
    res.data.valueRanges[0] &&
    res.data.valueRanges[0].values
      ? res.data.valueRanges[0].values[0]
      : [];
  const existingUrlValues =
    res.data.valueRanges &&
    res.data.valueRanges[1] &&
    res.data.valueRanges[1].values
      ? res.data.valueRanges[1].values[0]
      : [];
  const maxLen = Math.max(promptValues.length, existingUrlValues.length);

  const rows = [];
  for (let i = 0; i < maxLen; i++) {
    const prompt = promptValues[i] ? String(promptValues[i]).trim() : "";
    const existingUrl = existingUrlValues[i]
      ? String(existingUrlValues[i]).trim()
      : "";
    rows.push({ prompt, existingUrl, row: START_ROW + i });
  }
  return rows;
}

async function writeUrlToSheet(sheets, row, url) {
  const range = `${TAB_NAME}!D${row}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [[url]] },
  });
}
// waits for U1 button to appear
// uses an infinite loop to keep getting the last element and check if the U1 button appeared
async function waitForUpscaleButtons(page) {
  try {
    const startTime = Date.now();
    const timeout = 60 * 2000;

    while (true) {
      await sleepTwoSecs();
      const lastMessage = await getLastMessage(page);

      try {
        if (lastMessage.includes("U1")) {
          console.log("FOUND U1 BUTTON");
          await selectUpscaleButton(page);
          await getImageLinks(page);
          break;
        }
      } catch (error) {
        console.log("---------------ERROR---------------");
        console.log(error);
      }

      if (Date.now() - startTime > timeout) {
        console.log(
          "Timeout reached. Exiting loop as U1 button was not found."
        );
        imageLinksArray.push("");
        break;
      }
    }
  } catch (error) {
    console.log("---------------ERROR---------------");
    console.log(error);
  }
}

async function getLastMessage(page) {
  try {
    const lastMessage = page.locator("li.messageListItem__5126c").last();
    const lastMessageText = await lastMessage.innerText();
    return lastMessageText;
  } catch (error) {
    console.log("----------------ERROR------------------");
    console.log(error);
  }
}

async function selectUpscaleButton(page) {
  try {
    const upscaleButton = page
      .locator(`button:has-text("U1")`)
      .locator("nth=-1");

    if (upscaleButton) {
      await upscaleButton.click();
    }
  } catch (error) {
    console.log("----------------ERROR------------------");
    console.log(error);
  }
}
// a recursive function that gets the upscaled image after the U1 button has been clicked
async function getImageLinks(page) {
  try {
    const lastMessageText = await getLastMessage(page);
    if (
      lastMessageText.includes("Vary (Strong)") &&
      lastMessageText.includes("Web")
    ) {
      await sleepTwoSecs();
      const linkLocator = page.locator("a.originalLink_af017a").last();

      const lastImageLink = await linkLocator.getAttribute("data-safe-src");

      const ImageURL = lastImageLink.replace(/&width=\d+&height=\d+/, "");

      imageLinksArray.push(ImageURL);
    } else await getImageLinks(page);
  } catch (error) {
    console.log(error);
  }
}

// function for inserting images into cells in google sheets
async function insertImagesToCell(
  page,
  insertMenuButton,
  imageButton,
  insertImageToCellButton,
  link,
  index
) {
  if (index == 0) {
    await page.waitForTimeout(5000);
    await page.keyboard.press("ArrowDown", { delay: 300 });
    await page.keyboard.press("ArrowRight", { delay: 300 });
    await page.keyboard.press("ArrowRight", { delay: 300 });
    await page.keyboard.press("ArrowRight", { delay: 300 });
    await page.keyboard.press("ArrowRight", { delay: 300 });
    if (link == "") return;
  } else {
    await page.waitForTimeout(1000);
    await page.keyboard.press("ArrowDown", { delay: 300 });
  }

  if (link == "") return await page.keyboard.press("ArrowDown", { delay: 300 });

  await insertMenuButton.click();
  await imageButton.hover();
  await insertImageToCellButton.click();
  await page.waitForTimeout(1000);
  await page.keyboard.press("Tab", { delay: 300 });
  await page.keyboard.press("Tab", { delay: 300 });
  await page.keyboard.press("ArrowRight", { delay: 300 });
  await page.keyboard.press("ArrowRight", { delay: 300 });
  await page.keyboard.press("Enter", { delay: 300 });
  await page.keyboard.press("Tab", { delay: 300 });
  await page.keyboard.insertText(link);
  await page.waitForTimeout(3000);
  await sleepTwoSecs();
  await page.keyboard.press("Tab", { delay: 300 });
  await page.keyboard.press("Tab", { delay: 300 });
  await page.keyboard.press("Enter", { delay: 300 });
}

/* ========== MAIN FLOW ========== */
(async () => {
  let sheets;
  try {
    sheets = await sheetsClient();
  } catch (err) {
    console.error(
      "Google Sheets auth error:",
      err && err.message ? err.message : err
    );
    process.exit(1);
  }

  const rows = await readSheetPrompts(sheets);
  if (!rows || rows.length === 0) {
    console.log(
      "No prompts found in",
      `${TAB_NAME}!${PROMPT_COLUMN}${START_ROW}:${PROMPT_COLUMN}`
    );
    process.exit(0);
  }
  console.log(
    `Found ${rows.length} rows. Will process rows where column D is empty.`
  );

  const browser = await chromium.connectOverCDP("http://localhost:9222");

  const chromeContext = browser.contexts()[0];
  const discordPage = chromeContext.pages()[0];

  if (CHANNEL_URL) {
    console.log("Opening channel:", CHANNEL_URL);
    await discordPage.goto(CHANNEL_URL, { waitUntil: "domcontentloaded" });
  } else {
    console.log(
      "No CHANNEL_URL set. Opening discord.com/app — please navigate manually to your Midjourney channel."
    );
    await discordPage.goto("https://discord.com/app", {
      waitUntil: "domcontentloaded",
    });
  }

  const textFieldSelector = 'div[role="textbox"]';

  const textField = discordPage.locator(textFieldSelector);
  let indexForURLS = 0;

  for (const item of rows) {
    const { prompt, existingUrl, row } = item;

    if (existingUrl) {
      console.log(`Row ${row}: already has URL in column D — skipping.`);
      imageLinksArray.push(existingUrl);
      indexForURLS++;
      continue;
    }

    try {
      // console.log(`Row ${row}: preparing to send prompt -> ${prompt}`);

      await textField.click();
      await textField.fill(prompt, { timeout: 1000 });
      await discordPage.waitForTimeout(1000);

      await discordPage.keyboard.press("Enter");

      await waitForUpscaleButtons(discordPage);
      await writeUrlToSheet(sheets, row, imageLinksArray[indexForURLS]);

      console.log(imageLinksArray);

      indexForURLS++;
    } catch (err) {
      console.error(`Row ${row}: error:`, err && err.stack ? err.stack : err);
      continue;
    }
  } // end loop
  const sheetsPage = await chromeContext.newPage();

  await discordPage.close();

  await sheetsPage.goto(
    "https://docs.google.com/spreadsheets/d/1ibOWAjmOZMeEBtetqKq6-0wmxDM0eHWj8EsLSAu0Xc0",
    { waitUntil: "domcontentloaded" }
  );

  const insertButton = sheetsPage.getByRole("menuitem", { name: "Insert" });
  const imageButton = sheetsPage.locator('[aria-label="Image g"]');
  const insertImageInCellButton = sheetsPage.locator(
    '[aria-label="Insert image in cell i"]'
  );

  let indexForInserting = 0;
  for (const link of imageLinksArray) {
    await insertImagesToCell(
      sheetsPage,
      insertButton,
      imageButton,
      insertImageInCellButton,
      link,
      indexForInserting
    );
    indexForInserting++;
  }

  await browser.close();
  console.log("All done. Browser closed.");
})();
