/**
 * ============================================
 * 台湾证券交易所（TWSE）权证数据下载脚本
 * ============================================
 * 
 * 功能：从 TWSE 服务器获取权证（warrant）股票信息，
 *       转换为 CSV 格式，并保存为 Excel 兼容文件
 * 
 * 架构概览：
 * 1. 导入依赖 → 2. 设置配置 → 3. 定义工具函数 → 4. 主程序流程
 */

/**
 * 使用方法：
 *  - 默认：node download-twse-warrant-all.mjs
 *    → 产出文件：warrant-all-<YYYYMMDD>.csv
 *  - 自定义输出：node download-twse-warrant-all.mjs myfile.csv
 */

// ============ 第一部分：导入 Node.js 内置模块 ============

// 文件操作：writeFile 用来写入文件内容到磁盘
import { writeFile, rename, unlink, mkdir } from 'node:fs/promises';

// 路径处理：处理文件路径、目录名等
import { dirname, join, resolve } from 'node:path';

// URL 转换：把 import.meta.url 转成本地文件路径
import { fileURLToPath } from 'node:url';

// ============ 第二部分：全局配置常量 ============

// 数据源 URL：TWSE 的权证股票 API 端点（返回 JSON 格式）
const SOURCE_URL = 'https://www.twse.com.tw/rwd/zh/stock/warrantStock?response=json';

// 来源网页 URL：用于 HTTP 请求头中的 Referer（让服务器相信请求来自网页）
const REFERER_URL = 'https://www.twse.com.tw/zh/products/securities/warrant/infomation/stock.html';

// 脚本所在目录：用来生成输出文件的基础路径
const SCRIPT_DIR = dirname(fileURLToPath(import.meta.url));

// UTF-8 字节顺序标记：Excel 用来识别 UTF-8 编码的特殊字符（\uFEFF）
const UTF8_BOM = '\uFEFF';

// 网络请求最大重试次数：如果连接失败，会自动重试多次
const MAX_ATTEMPTS = 3;

// fetch 超时（毫秒）
const FETCH_TIMEOUT_MS = 10000; // 10 秒

// 重试延迟基底（毫秒），实际延迟会乘以尝试次数
const RETRY_DELAY_MS = 3000;

// ============ 第三部分：工具函数（功能模块） ============

/**
 * 函数：生成台湾时区的日期戳
 * 输入：date（可选，默认使用当前时间）
 * 输出：例如 "20260530"（年月日格式）
 * 用途：用来生成带日期的文件名
 */
function taiwanDateStamp(date = new Date()) {
  // 使用台湾时区（Asia/Taipei）格式化日期为 "YYYY-MM-DD"
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date);
  
  // 把格式化的结果转成对象，便于提取年月日
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  
  // 拼接成 "YYYYMMDD" 格式（不带分隔符）
  return `${values.year}${values.month}${values.day}`;
}

/**
 * 函数：根据命令行参数决定输出文件路径
 * 输入：process.argv[2]（命令行第二个参数，如果有的话）
 * 输出：完整的文件路径
 * 用途：允许用户自定义输出文件名，或使用默认名称
 */
function outputPathFromArgs() {
  // 默认文件名：warrant-all-日期.csv（例如 warrant-all-20260530.csv）
  const fallback = `warrant-all-${taiwanDateStamp()}.csv`;
  
  // 如果用户提供了第二个参数，就用用户的；否则使用默认值
  let requested = process.argv[2] || fallback;
  // 确保扩展名为 .csv
  if (!/\.csv$/i.test(requested)) {
    requested = `${requested}.csv`;
  }
  
  // 把相对路径转成绝对路径
  return resolve(SCRIPT_DIR, requested);
}

// 最终的输出文件路径（会被用在下面）
const OUTPUT_FILE = outputPathFromArgs();

// 临时文件路径（为了安全，先写入临时文件，然后再改名为正式文件）
const TEMP_OUTPUT_FILE = `${OUTPUT_FILE}.tmp`;

/**
 * 函数：为 CSV 格式转义单元格内容
 * 输入：任意值（数字、字符串、null 等）
 * 输出：符合 CSV 标准的字符串（加引号、转义等）
 * 用途：处理特殊字符，使 Excel 能正确识别
 */
function escapeCsv(value) {
  // 把值转成字符串
  const text = String(value ?? '')
    // 移除 HTML 标签（如 <br>、<span> 等）
    .replace(/<[^>]*>/g, '')
    // 把 HTML 空格转成普通空格
    .replace(/&nbsp;/g, ' ')
    // 去掉前后的空白
    .trim();
  
  // 如果文本包含特殊字符（引号、逗号、换行），就用双引号包裹
  // 并把内部的双引号转义（变成两个双引号）
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : `"${text}"`;
}

/**
 * 函数：为 Excel 编码单元格内容
 * 输入：任意值
 * 输出：适合 Excel 显示的格式
 * 用途：防止 Excel 把某些值误解为公式或数字
 * 
 * 例如：
 * - "0123" 会被 Excel 识别为数字 123（丢失前导 0）
 * - "ABC" 可能被识别为公式
 * 所以用 ='0123' 或 ="ABC" 强制让 Excel 把它们当成文本
 */
function codeForExcel(value) {
  const text = String(value ?? '');
  
  // 如果文本以 0 开头，或者以字母结尾，需要用公式格式
  return /^0/.test(text) || /[A-Za-z]$/.test(text) ? `=${JSON.stringify(text)}` : text;
}

/**
 * 函数：构建分组标题行
 * 输入：
 *   - fields：列名数组
 *   - groups：分组信息数组（每个元素包含 start 和 title）
 * 输出：包含分组标题的行数组
 * 用途：在 CSV 的第二行添加分组标题，用来组织相关列
 */
function buildGroupRow(fields, groups) {
  // 创建一个和列数一样长的空行
  const row = Array(fields.length).fill('');
  
  // 在指定位置填入分组标题
  for (const group of groups ?? []) {
    row[group.start] = group.title;
  }
  
  return row;
}

// 确保输出文件所在目录存在
async function ensureDirectoryExists(filePath) {
  const dir = dirname(filePath);
  try {
    await mkdir(dir, { recursive: true });
  } catch (err) {
    // 如果创建目录失败，抛出错误
    throw new Error(`Failed to create directory ${dir}: ${err && err.message ? err.message : err}`);
  }
}


// ============ 第四部分：主程序流程 ============

/**
 * 异步函数：主程序入口
 * 
 * 执行步骤：
 * 1. 从 TWSE 服务器获取权证数据（带重试机制）
 * 2. 校验响应数据是否有效
 * 3. 将数据转换为 CSV 格式（带 Excel 兼容处理）
 * 4. 保存到文件
 * 5. 打印执行结果
 */
async function main() {
  // ===== 步骤 0：預取 Referer 頁面（模擬瀏覽器先載入網頁） =====
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

  // ===== 步骤 1：网络请求（带自动重试机制） =====

  let response;
  let lastError;
  
  // 尝试最多 3 次请求
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt += 1) {
    try {
      console.error(`Attempt ${attempt}/${MAX_ATTEMPTS} fetching JSON...`);
      // 为每次请求建立超时控制（避免长时间挂起）
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);

      try {
        response = await fetch(SOURCE_URL, {
          signal: controller.signal,
          headers: {
            accept: 'application/json, text/plain, */*',
            // Referer 头：告诉服务器这个请求来自网页（某些网站会检查这个）
            referer: REFERER_URL,
            // User-Agent：伪装成浏览器请求（有些服务器会拒绝爬虫）
            'user-agent': 'Mozilla/5.0',
          },
        });
      } finally {
        clearTimeout(timeoutId);
      }

      console.error('Fetch JSON: OK');
      // 请求成功就跳出循环
      break;
    } catch (error) {
      // 记录错误
      lastError = error;

      // 如果这是超时，给出更友善的错误信息
      if (error && error.name === 'AbortError') {
        lastError = new Error(`Request timed out after ${FETCH_TIMEOUT_MS}ms`);
      }

      // 如果还有重试机会，等一段时间再试（避免频繁请求导致被封 IP）
      if (attempt < MAX_ATTEMPTS) {
        const waitTime = attempt * RETRY_DELAY_MS;
        console.error(`Retrying after ${waitTime}ms...`);
        await new Promise((resolve) => setTimeout(resolve, waitTime));
      }
    }
  }

  // 检查是否成功获取响应
  if (!response && lastError) {
    // 如果所有重试都失败了，抛出最后一个错误
    throw lastError;
  }

  // ===== 步骤 2：校验 HTTP 响应状态 =====
  
  if (!response.ok) {
    // HTTP 状态码不是 2xx（例如 404、500 等）
    throw new Error(`TWSE returned HTTP ${response.status}`);
  }

  // ===== 步骤 3：解析 JSON 并校验数据结构 =====
  
  const payload = await response.json();
  
  // 检查响应是否符合预期格式
  if (payload.stat !== 'OK' || !Array.isArray(payload.fields) || !Array.isArray(payload.data)) {
    throw new Error(`Unexpected TWSE response: ${JSON.stringify({ stat: payload.stat })}`);
  }

  // ===== 步骤 4：构造 CSV 内容 =====
  
  // CSV 应该包含以下行：
  // 1. 标题行（例如 "台湾权证股票信息"）
  // 2. 分组标题行（用来组织列）
  // 3. 列标题行（字段名）
  // 4. 数据行（实际的权证信息）
  
  const csvRows = [
    [payload.title],                                    // 第 1 行：标题
    buildGroupRow(payload.fields, payload.groups),      // 第 2 行：分组标题
    payload.fields,                                      // 第 3 行：列标题
    // 第 4 行及以后：数据行
    // 注意：第一列（股票代码）需要特殊处理（codeForExcel），防止被 Excel 识别为数字
    ...payload.data.map((row) => 
      row.map((value, index) => 
        index === 0 ? codeForExcel(value) : value
      )
    ),
  ];

  // ===== 步骤 5：将 CSV 行转换为字符串 =====
  
  // 每行的每个单元格用 escapeCsv 处理（转义特殊字符）
  // 同一行的单元格用逗号分隔
  // 不同行用 Windows 换行符 \r\n 分隔
  const csv = csvRows
    .map((row) => 
      row.map(escapeCsv).join(',')  // 处理每个单元格，然后用逗号连接
    )
    .join('\r\n');                  // 用换行符连接所有行

  // ===== 步骤 6：保存文件 =====
  
  // 先写临时文件（如果写入失败，不会覆盖正式文件）
  // UTF8_BOM 是一个特殊字符，告诉 Excel "这是 UTF-8 编码"
  // 确保输出目录存在
  await ensureDirectoryExists(OUTPUT_FILE);

  // 先写入临时文件，再将其改名为正式文件（原子性更好，避免中间状态）
  await writeFile(TEMP_OUTPUT_FILE, UTF8_BOM + csv + '\r\n', 'utf8');
  try {
    // 如果目标文件已存在，尝试先删除（避免 Windows 上 rename 失败）
    await unlink(OUTPUT_FILE);
  } catch (_) {
    // 忽略不存在或删除错误
  }
  await rename(TEMP_OUTPUT_FILE, OUTPUT_FILE);

  // ===== 步骤 7：打印执行结果 =====
  
  console.log(JSON.stringify({
    output: OUTPUT_FILE,        // 输出文件的完整路径
    title: payload.title,       // 数据的标题
    date: payload.date,         // 数据的日期
    dataRows: payload.data.length,  // 有多少行权证数据
    csvLines: csvRows.length,       // CSV 总共有多少行（包括标题）
    encoding: 'UTF-8 with BOM',     // 编码方式
  }, null, 2));
}

// ===== 步骤 8：执行主程序（处理错误） =====

// 调用主函数，如果出错就打印错误信息，并设置进程退出码为 1
main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
