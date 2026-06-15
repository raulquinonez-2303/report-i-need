const fs = require("node:fs/promises");
const path = require("node:path");
const { chromium } = require("@playwright/test");
const { google } = require("googleapis");

const TARGET_URL = "https://www.ineedtours.com/es/tours.html";
const SPREADSHEET_ID = "1wBZKRRFBJZAUWdPsGa9hX3Ifak8zjKktaTdHFsx3Rms";
const SHEET_NAME = "Hoja 1";
const OUTPUT_DIR = path.resolve("artifacts");

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Falta la variable de entorno requerida: ${name}`);
  }
  return value;
}

async function ensureDirs() {
  await fs.mkdir(path.join(OUTPUT_DIR, "videos"), { recursive: true });
  await fs.mkdir(path.join(OUTPUT_DIR, "screenshots"), { recursive: true });
}

async function writeRunMetadata(metadata) {
  const destination = path.join(OUTPUT_DIR, "run-metadata.json");
  await fs.writeFile(destination, JSON.stringify(metadata, null, 2), "utf8");
}

async function appendRunLogToSheets(status) {
  const credentialsRaw = requiredEnv("CREDENCIALES_JSON");
  const credentials = JSON.parse(credentialsRaw);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  const sheets = google.sheets({ version: "v4", auth });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:E`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [
        [
          new Date().toISOString(),
          status,
          "login+video",
          process.env.GITHUB_RUN_ID || "local",
          "Pendiente: descargar y pegar reporte final"
        ]
      ]
    }
  });
}

async function run() {
  const user = requiredEnv("USER_INEED");
  const pass = requiredEnv("PASS_INEED");

  await ensureDirs();

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    recordVideo: {
      dir: path.join(OUTPUT_DIR, "videos"),
      size: { width: 1366, height: 768 }
    },
    viewport: { width: 1366, height: 768 }
  });
  const page = await context.newPage();
  const pageVideo = page.video();

  let loginStatus = "LOGIN_OK";

  try {
    await page.goto(TARGET_URL, { waitUntil: "domcontentloaded", timeout: 120000 });
    await page.click("#ContentPlaceHolder_ctl00_ctl00_ctl06_ctl00_lnkHeaderMiCuenta");
    await page.waitForSelector("#ctl26_ctl00_txtUsuario", { state: "visible", timeout: 30000 });

    await page.fill("#ctl26_ctl00_txtUsuario", user);
    await page.fill("#ctl26_ctl00_txtClave", pass);
    await page.click("#ctl26_ctl00_btnLogin");

    await page.waitForTimeout(7000);
    await page.screenshot({
      path: path.join(OUTPUT_DIR, "screenshots", "post-login.png"),
      fullPage: true
    });
  } catch (error) {
    loginStatus = "LOGIN_ERROR";
    await page.screenshot({
      path: path.join(OUTPUT_DIR, "screenshots", "login-error.png"),
      fullPage: true
    });
    throw error;
  } finally {
    await context.storageState({ path: path.join(OUTPUT_DIR, "storage-state.json") });
    await context.close();
    await browser.close();
  }

  if (!pageVideo) {
    throw new Error("No se pudo obtener el video de la sesión.");
  }

  const rawVideoPath = await pageVideo.path();
  const finalVideoPath = path.join(OUTPUT_DIR, "videos", "login-session.webm");
  if (rawVideoPath !== finalVideoPath) {
    await fs.copyFile(rawVideoPath, finalVideoPath);
  }

  await writeRunMetadata({
    loginStatus,
    runAt: new Date().toISOString(),
    targetUrl: TARGET_URL,
    videoFile: path.relative(process.cwd(), finalVideoPath).replaceAll("\\", "/")
  });

  await appendRunLogToSheets(loginStatus);
}

run().catch(async (error) => {
  console.error("Fallo en ejecución del bot:", error.message);
  process.exitCode = 1;
});
