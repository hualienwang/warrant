import { writeFile, rename, unlink, mkdir } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const SOURCE_URL = 'https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?response=json&type=ALLBUT0999&date=';
const REFERER_URL = 'https://www.twse.com.tw/zh/page/trading/exchange/MI_INDEX.html';
const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));
const UTF8_BOM = '\uFEFF';
const MAX_ATTEMPTS = 3;
const FETCH_TIMEOUT_MS = 30000;
const RETRY_DELAY_MS = 3000;
const DEFAULT_WORKBOOK_FILE = join(SCRIPT_DIR, 'Excel股票處理.xlsx');

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

function parseArgs(argv = process.argv.slice(2)) {
  const options = {
    output: null,
    refreshExcel: false,
    workbook: DEFAULT_WORKBOOK_FILE,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === '--refresh-excel') {
      options.refreshExcel = true;
      const next = argv[index + 1];
      if (next && !next.startsWith('--') && !/\.csv$/i.test(next)) {
        options.workbook = resolve(SCRIPT_DIR, next);
        index += 1;
      }
      continue;
    }

    if (arg === '--workbook') {
      const next = argv[index + 1];
      if (!next) {
        throw new Error('Missing workbook path after --workbook');
      }
      options.workbook = resolve(SCRIPT_DIR, next);
      index += 1;
      continue;
    }

    if (arg.startsWith('--')) {
      throw new Error(`Unknown option: ${arg}`);
    }

    if (!options.output) {
      options.output = arg;
      continue;
    }

    throw new Error(`Unexpected argument: ${arg}`);
  }

  return options;
}

const OPTIONS = parseArgs();

function outputPathFromArgs() {
  const fallback = `twse-stocks-${taiwanDateStamp()}.csv`;
  let requested = OPTIONS.output || fallback;
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

function execFileAsync(file, args, options = {}) {
  return new Promise((resolvePromise, rejectPromise) => {
    execFile(file, args, options, (error, stdout, stderr) => {
      if (error) {
        error.stdout = stdout;
        error.stderr = stderr;
        rejectPromise(error);
        return;
      }
      resolvePromise({ stdout, stderr });
    });
  });
}

async function refreshExcelWorkbook(workbookPath) {
  if (process.platform !== 'win32') {
    throw new Error('Excel refresh requires Windows with desktop Excel installed.');
  }

  const escapedWorkbookPath = workbookPath.replace(/'/g, "''");
  const script = String.raw`
$ErrorActionPreference = 'Stop'
$workbookPath = '${escapedWorkbookPath}'
$excel = $null
$workbook = $null

try {
  function Invoke-ComRetry {
    param(
      [string]$Description,
      [scriptblock]$Action,
      [int]$Attempts = 60,
      [int]$DelayMilliseconds = 1000
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt += 1) {
      try {
        return & $Action
      }
      catch {
        if ($attempt -ge $Attempts) {
          throw
        }
        Write-Output "[Excel] $Description is busy; retry $attempt/$Attempts..."
        Start-Sleep -Milliseconds $DelayMilliseconds
      }
    }
  }

  Write-Output "[Excel] resolving workbook path..."
  $resolvedWorkbookPath = (Resolve-Path -LiteralPath $workbookPath).Path
  try {
    $lockTest = [System.IO.File]::Open($resolvedWorkbookPath, 'Open', 'ReadWrite', 'None')
    $lockTest.Close()
  }
  catch {
    throw "Workbook is locked. Close Excel股票處理.xlsx and any hidden EXCEL.EXE processes, then run again. Path: $resolvedWorkbookPath"
  }

  Write-Output "[Excel] starting Excel..."
  $excel = Invoke-ComRetry "start Excel" { New-Object -ComObject Excel.Application } 10 1000
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false

  Write-Output "[Excel] opening workbook: $resolvedWorkbookPath"
  $workbook = Invoke-ComRetry "open workbook" { $excel.Workbooks.Open($resolvedWorkbookPath) } 30 1000

  Write-Output "[Excel] disabling background refresh..."
  foreach ($connection in $workbook.Connections) {
    try { $connection.OLEDBConnection.BackgroundQuery = $false } catch {}
    try { $connection.ODBCConnection.BackgroundQuery = $false } catch {}
  }

  foreach ($worksheet in $workbook.Worksheets) {
    foreach ($queryTable in $worksheet.QueryTables) {
      try { $queryTable.BackgroundQuery = $false } catch {}
    }
    foreach ($listObject in $worksheet.ListObjects) {
      try { $listObject.QueryTable.BackgroundQuery = $false } catch {}
    }
  }

  Write-Output "[Excel] refreshing queries..."
  Invoke-ComRetry "refresh queries" { $workbook.RefreshAll() } 30 1000 | Out-Null
  Write-Output "[Excel] waiting 30 seconds for query refresh..."
  Start-Sleep -Seconds 30
  for ($attempt = 1; $attempt -le 30; $attempt += 1) {
    try {
      if ($excel.Ready) {
        break
      }
    }
    catch {}
    Write-Output "[Excel] Excel is still busy; wait $attempt/30..."
    Start-Sleep -Seconds 1
  }
  Write-Output "[Excel] saving workbook..."
  Invoke-ComRetry "save workbook" { $workbook.Save() } 30 1000 | Out-Null
  Write-Output "Excel refreshed and saved: $resolvedWorkbookPath"
}
finally {
  if ($workbook -ne $null) {
    try { $workbook.Close($true) } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
  }
  if ($excel -ne $null) {
    try { $excel.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
`;

  const powershell = 'powershell.exe';
  return execFileAsync(
    powershell,
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', script],
    { windowsHide: true, timeout: 5 * 60 * 1000 },
  );
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
  const finalOutput = OPTIONS.output
    ? outputPathFromArgs()
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

  if (OPTIONS.refreshExcel) {
    console.error(`Refreshing Excel workbook: ${OPTIONS.workbook}`);
    const { stdout, stderr } = await refreshExcelWorkbook(OPTIONS.workbook);
    if (stderr) {
      console.error(stderr.trim());
    }
    if (stdout) {
      console.error(stdout.trim());
    }
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
