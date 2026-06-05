import { writeFile, rename, unlink, mkdir } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const SOURCE_URL = 'https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?response=json&type=ALLBUT0999&date=';
const REFERER_URL = 'https://www.twse.com.tw/zh/page/trading/exchange/MI_INDEX.html';
const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const UTF8_BOM = '\uFEFF';
const MAX_ATTEMPTS = 3;
const FETCH_TIMEOUT_MS = 30000;
const RETRY_DELAY_MS = 3000;

function taiwanDateStamp(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${values.year}${values.month}${values.day}`;
}

function outputPathFromArgs() {
  const fallback = `twse-stocks-${taiwanDateStamp()}.csv`;
  let requested = process.argv[2] || fallback;
  if (!/\.csv$/i.test(requested)) {
    requested = `${requested}.csv`;
  }
  return resolve(SCRIPT_DIR, requested);
}

// (output path is determined inside main() to support auto date fallback)

function escapeCsv(value) {
  const text = String(value ?? '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .trim();
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : `"${text}"`;
}

function codeForExcel(value) {
  const text = String(value ?? '');
  return /^0/.test(text) || /[A-Za-z]$/.test(text) ? `=${JSON.stringify(text)}` : text;
}

function buildGroupRow(fields, groups) {
  const row = Array(fields.length).fill('');
  for (const group of groups ?? []) {
    row[group.start] = group.title;
  }
  return row;
}

async function ensureDirectoryExists(filePath) {
  const dir = dirname(filePath);
  try {
    await mkdir(dir, { recursive: true });
  } catch (err) {
    throw new Error(`Failed to create directory ${dir}: ${err && err.message ? err.message : err}`);
  }
}

async function fetchWithRetry(url) {
  let response;
  let lastError;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt += 1) {
    try {
      console.error(`Attempt ${attempt}/${MAX_ATTEMPTS} fetching ...`);
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

      try {
        response = await fetch(url, {
          signal: controller.signal,
          headers: {
            accept: 'application/json, text/plain, */*',
            referer: REFERER_URL,
            'user-agent': 'Mozilla/5.0',
          },
        });
      } finally {
        clearTimeout(timeoutId);
      }

      console.error('Fetch JSON: OK');
      break;
    } catch (error) {
      lastError = error;
      if (error && error.name === 'AbortError') {
        lastError = new Error(`Request timed out after ${FETCH_TIMEOUT_MS}ms`);
      }
      if (attempt < MAX_ATTEMPTS) {
        const waitTime = attempt * RETRY_DELAY_MS;
        console.error(`Retrying after ${waitTime}ms...`);
        await new Promise((resolve) => setTimeout(resolve, waitTime));
      }
    }
  }

  if (!response && lastError) {
    throw lastError;
  }
  if (!response.ok) {
    throw new Error(`TWSE returned HTTP ${response.status}`);
  }
  return response;
}

async function main() {
  let dateStr = process.argv[3] || taiwanDateStamp();

  try {
    const controllerRef = new AbortController();
    const refTimeout = setTimeout(() => controllerRef.abort(), 5000);
    const refResp = await fetch(REFERER_URL, {
      signal: controllerRef.signal,
      headers: {
        accept: 'text/html,application/xhtml+xml',
        'user-agent': 'Mozilla/5.0',
      },
    });
    try { await refResp.text(); } catch (_) {}
    clearTimeout(refTimeout);
    console.error('Prefetch referer: OK');
  } catch (err) {
    console.error('Warning: referer prefetch failed:', err && err.message ? err.message : err);
  }

  let payload;
  let usedDate = dateStr;

  // try up to 7 previous days (for weekends/holidays)
  for (let day = 0; day < 7; day++) {
    const d = new Date();
    d.setDate(d.getDate() - day);
    if (process.argv[3]) {
      // if user specified a date, parse it
      const y = parseInt(dateStr.slice(0,4), 10);
      const m = parseInt(dateStr.slice(4,6), 10) - 1;
      const dayNum = parseInt(dateStr.slice(6,8), 10);
      d.setFullYear(y, m, dayNum - day);
    }
    const tryDate = taiwanDateStamp(d);
    if (day > 0) {
      console.error(`No data for ${usedDate}, trying ${tryDate}...`);
    }
    usedDate = tryDate;
    const url = SOURCE_URL + tryDate;

    const response = await fetchWithRetry(url);
    payload = await response.json();

    if (payload.stat === 'OK' && Array.isArray(payload.tables)) {
      break;
    }
    if (day === 6) {
      throw new Error(`No trading data found for the past 7 days (last tried ${tryDate})`);
    }
  }

  // Find the table with stock data (has "證券代號" field)
  const stockTable = payload.tables.find(t =>
    Array.isArray(t.fields) && t.fields.includes('證券代號') && t.data.length > 0
  );
  if (!stockTable) {
    throw new Error('Stock data table not found in response (no table with 證券代號)');
  }

  // If we fell back to a different date, adjust the output filename
  const finalOutput = process.argv[2]
    ? resolve(SCRIPT_DIR, process.argv[2])
    : resolve(SCRIPT_DIR, `twse-stocks-${usedDate}.csv`);

  const csvRows = [
    [stockTable.title],
    buildGroupRow(stockTable.fields, stockTable.groups),
    stockTable.fields,
    ...stockTable.data.map((row) =>
      row.map((value, index) =>
        index === 0 ? codeForExcel(value) : value
      )
    ),
  ];

  const csv = csvRows
    .map((row) =>
      row.map(escapeCsv).join(',')
    )
    .join('\r\n');

  const tmpFile = `${finalOutput}.tmp`;
  await ensureDirectoryExists(finalOutput);
  await writeFile(tmpFile, UTF8_BOM + csv + '\r\n', 'utf8');
  try {
    await unlink(finalOutput);
  } catch (_) {}
  await rename(tmpFile, finalOutput);

  console.log(JSON.stringify({
    output: finalOutput,
    title: stockTable.title,
    date: payload.date,
    dataRows: stockTable.data.length,
    csvLines: csvRows.length,
    fields: stockTable.fields,
    encoding: 'UTF-8 with BOM',
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
