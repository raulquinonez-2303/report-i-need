const fs = require("node:fs/promises");
const path = require("node:path");
const { chromium } = require("@playwright/test");
const { google } = require("googleapis");
const XLSX = require("xlsx");

const TARGET_URL = "https://www.ineedtours.com/es/tours.html";
const BOOKINGS_URL = "https://www.ineedtours.com/V05/paginas/privadas/listado_reservas.aspx";
const SPREADSHEET_ID = "1wBZKRRFBJZAUWdPsGa9hX3Ifak8zjKktaTdHFsx3Rms";
const SHEET_NAME = "Reservas I Need Tours";
const OUTPUT_DIR = path.resolve("artifacts");
const EXCEL_DOWNLOAD_PATH = path.join(OUTPUT_DIR, "downloads", "reservas.xlsx");
const EXCEL_BUTTON_SELECTOR = "#ContentPlaceHolder_ctl02_lnkExcel";

function requiredEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Falta la variable de entorno requerida: ${name}`);
  }
  return value;
}

function quoteSheetName(sheetName) {
  return `'${sheetName.replaceAll("'", "''")}'`;
}

function sheetRange(a1Range) {
  return `${quoteSheetName(SHEET_NAME)}!${a1Range}`;
}

async function ensureDirs() {
  await fs.mkdir(path.join(OUTPUT_DIR, "videos"), { recursive: true });
  await fs.mkdir(path.join(OUTPUT_DIR, "screenshots"), { recursive: true });
  await fs.mkdir(path.join(OUTPUT_DIR, "downloads"), { recursive: true });
}

async function writeRunMetadata(metadata) {
  const destination = path.join(OUTPUT_DIR, "run-metadata.json");
  await fs.writeFile(destination, JSON.stringify(metadata, null, 2), "utf8");
}

function getSheetsClient() {
  const credentialsRaw = requiredEnv("CREDENCIALES_JSON");
  const credentials = JSON.parse(credentialsRaw);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  return google.sheets({ version: "v4", auth });
}

function excelDateToIso(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) {
      return "";
    }
    const asUtc = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
    return asUtc.toISOString().slice(0, 10);
  }
  if (typeof value === "string" && value.trim()) {
    const normalized = value.trim().replaceAll("/", "-");
    const directDate = new Date(normalized);
    if (!Number.isNaN(directDate.getTime())) {
      return directDate.toISOString().slice(0, 10);
    }

    const ddmmyyyy = normalized.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
    if (ddmmyyyy) {
      const day = Number(ddmmyyyy[1]);
      const month = Number(ddmmyyyy[2]);
      const year = Number(ddmmyyyy[3]);
      const asUtc = new Date(Date.UTC(year, month - 1, day));
      return asUtc.toISOString().slice(0, 10);
    }
  }
  return "";
}

function subtractDaysFromIso(isoDate, days) {
  if (!isoDate) {
    return "";
  }
  const date = new Date(`${isoDate}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  date.setUTCDate(date.getUTCDate() - days);
  return date.toISOString().slice(0, 10);
}

function formatIsoToDdMmYyyy(isoDate) {
  if (!isoDate) {
    return "";
  }
  const parts = isoDate.split("-");
  if (parts.length !== 3) {
    return "";
  }
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
}

function sanitizeMoney(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).replace("$", "").replaceAll(".", "").replace(",", ".").trim();
}

function transformReservationRows(excelPath) {
  const workbook = XLSX.readFile(excelPath, { cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("El archivo Excel descargado no tiene hojas.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const sourceRows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: ""
  });

  if (sourceRows.length <= 1) {
    return [];
  }

  return sourceRows.slice(1).map((row) => {
    const out = [...row];

    const travelDateIso = excelDateToIso(out[5]);
    out[5] = formatIsoToDdMmYyyy(travelDateIso);
    out[7] = formatIsoToDdMmYyyy(subtractDaysFromIso(travelDateIso, 15));

    out[9] = sanitizeMoney(out[9]);
    out[10] = sanitizeMoney(out[10]);
    return out;
  });
}

async function overwriteSheetWithRows(rows) {
  const sheets = getSheetsClient();
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetRange("A2:ZZ")
  });

  if (rows.length === 0) {
    return;
  }

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetRange("A2"),
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rows
    }
  });

  await applyDateFormatting(sheets);
}

async function applyDateFormatting(sheets) {
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets(properties(sheetId,title))"
  });
  const targetSheet = spreadsheet.data.sheets?.find(
    (sheet) => sheet.properties?.title === SHEET_NAME
  );

  if (!targetSheet?.properties?.sheetId && targetSheet?.properties?.sheetId !== 0) {
    throw new Error(`No se encontró la pestaña "${SHEET_NAME}" para aplicar formato de fecha.`);
  }

  const sheetId = targetSheet.properties.sheetId;
  const dateFormatRequest = {
    userEnteredFormat: {
      numberFormat: {
        type: "DATE",
        pattern: "dd/mm/yyyy"
      }
    }
  };

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              startColumnIndex: 5,
              endColumnIndex: 6
            },
            cell: dateFormatRequest,
            fields: "userEnteredFormat.numberFormat"
          }
        },
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: 1,
              startColumnIndex: 7,
              endColumnIndex: 8
            },
            cell: dateFormatRequest,
            fields: "userEnteredFormat.numberFormat"
          }
        }
      ]
    }
  });
}

async function goToBookingsPage(page) {
  const modalCloseButton = page.locator(
    "#modalbody button.close[data-dismiss='modal'][aria-hidden='true']"
  );
  const modalBody = page.locator("#modalbody");
  const modalBackdrop = page.locator(".modal-backdrop");

  // Close welcome modal with multiple strategies (button, outside click, Esc, JS fallback).
  if (await modalBody.isVisible({ timeout: 20000 }).catch(() => false)) {
    if (await modalCloseButton.isVisible({ timeout: 3000 }).catch(() => false)) {
      await modalCloseButton.click({ force: true }).catch(() => null);
    }

    if (await modalBody.isVisible({ timeout: 1500 }).catch(() => false)) {
      if (await modalBackdrop.isVisible({ timeout: 1000 }).catch(() => false)) {
        await modalBackdrop.click({ force: true }).catch(() => null);
      }
      await page.locator("body").click({ force: true, position: { x: 5, y: 5 } }).catch(() => null);
    }

    if (await modalBody.isVisible({ timeout: 1500 }).catch(() => false)) {
      await page.keyboard.press("Escape").catch(() => null);
    }

    if (await modalBody.isVisible({ timeout: 1500 }).catch(() => false)) {
      await page.evaluate(() => {
        const modal = document.querySelector("#modalbody");
        if (modal) {
          modal.classList.remove("in", "show");
          modal.setAttribute("aria-hidden", "true");
          modal.setAttribute("style", "display:none;");
        }
        document.querySelectorAll(".modal-backdrop").forEach((el) => el.remove());
        document.body.classList.remove("modal-open");
        document.body.style.removeProperty("padding-right");
      });
    }

    await modalBody.waitFor({ state: "hidden", timeout: 10000 }).catch(() => null);
  }

  // Then go directly to bookings page and continue with download flow.
  await page.goto(BOOKINGS_URL, { waitUntil: "domcontentloaded", timeout: 120000 });

  await page.waitForSelector(EXCEL_BUTTON_SELECTOR, { state: "visible", timeout: 60000 });
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
  let downloadedRows = 0;
  let videoFile = "";

  try {
    await page.goto(TARGET_URL, { waitUntil: "domcontentloaded", timeout: 120000 });
    await page.click("#ContentPlaceHolder_ctl00_ctl00_ctl06_ctl00_lnkHeaderMiCuenta");
    await page.waitForSelector("#ctl26_ctl00_txtUsuario", { state: "visible", timeout: 30000 });

    await page.fill("#ctl26_ctl00_txtUsuario", user);
    await page.fill("#ctl26_ctl00_txtClave", pass);
    await page.click("#ctl26_ctl00_btnLogin");

    await goToBookingsPage(page);

    const [download] = await Promise.all([
      page.waitForEvent("download", { timeout: 90000 }),
      page.click(EXCEL_BUTTON_SELECTOR)
    ]);
    await download.saveAs(EXCEL_DOWNLOAD_PATH);

    const transformedRows = transformReservationRows(EXCEL_DOWNLOAD_PATH);
    downloadedRows = transformedRows.length;
    await overwriteSheetWithRows(transformedRows);

    await page.waitForTimeout(2000);
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
  videoFile = path.relative(process.cwd(), finalVideoPath).replaceAll("\\", "/");

  await writeRunMetadata({
    loginStatus,
    runAt: new Date().toISOString(),
    targetUrl: TARGET_URL,
    bookingsUrl: BOOKINGS_URL,
    downloadedExcel: path.relative(process.cwd(), EXCEL_DOWNLOAD_PATH).replaceAll("\\", "/"),
    rowsPastedToSheet: downloadedRows,
    videoFile
  });
}

run().catch(async (error) => {
  console.error("Fallo en ejecución del bot:", error.message);
  process.exitCode = 1;
});
