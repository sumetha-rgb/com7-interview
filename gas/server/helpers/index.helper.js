function getRowsFromSheet(sheet) {
  // --- Get values form sheet
  const values = sheet.getDataRange().getValues();

  // --- Check if sheet is empty
  if (!values || values.length < 2) {
    return [];
  }

  // --- Get header
  const header = values[0];
  const rows = [];

  // --- Get rows
  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    if (rowData.join("").trim() === "") continue;

    const rowObj = {};

    // --- Assign values
    header.forEach((header, index) => {
      let rowValue = rowData[index];

      if (rowValue instanceof Date) {
        rowValue = Utilities.formatDate(
          rowValue,
          "Asia/Bangkok",
          "yyyy-MM-dd'T'HH:mm:ssXXX",
        );
      }

      rowObj[header] = rowValue;
    });

    rows.push(rowObj);
  }

  return rows;
}

// For merge map values to rows
const mergeMapValuesToRows = (values) => {
  if (!values || values.length < 2) return [];
  const header = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const rowData = values[i];
    if (rowData.join("").trim() === "") continue;
    const rowObj = {};
    header.forEach((h, colIndex) => {
      rowObj[h] = rowData[colIndex] !== undefined ? rowData[colIndex] : "";
    });
    rows.push(rowObj);
  }
  return rows;
};

function findRowIndex(sheetOrRows, rowValue, columnName) {
  // --- OPTIMIZATION: ถ้า caller ส่ง rows (array) ที่โหลดไว้แล้วมาให้ ใช้ตรงๆ
  // เลย ไม่อ่านชีทซ้ำ (ลดจำนวนครั้งที่เปิด/อ่าน Spreadsheet ต่อ 1 request)
  // ถ้า caller ยังส่ง Sheet object แบบเดิมมา ก็ยังทำงานเหมือนเดิมทุกประการ
  // (ผลลัพธ์ index ที่ได้เหมือนเดิมทุกกรณี ไม่กระทบ logic)
  const data = Array.isArray(sheetOrRows)
    ? sheetOrRows
    : getRowsFromSheet(sheetOrRows);
  const rowIndex = data.findIndex(
    (row) => String(row[columnName]) === String(rowValue),
  );
  return rowIndex;
}

function formatDateTh(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "";
  return Utilities.formatDate(dateObj, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
}

function formatTimeTh(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "";
  return Utilities.formatDate(dateObj, "Asia/Bangkok", "HH:mm");
}

function normalizeDate(date) {
  const d = new Date(date);

  const year = d.getFullYear();
  if (year > 2400) {
    d.setFullYear(year - 543);
  }
  return d;
}

function toDateOnly(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  d.setMilliseconds(0);
  return d;
}

const parseThaiDateTime = (dateString) => {
  if (!dateString) return new Date(0);

  try {
    const parts = dateString.trim().split(" ");
    const datePart = parts[0];
    const timePart = parts[1] || "00:00:00";

    const dateParts = datePart.split("/");
    if (dateParts.length !== 3) return new Date(dateString);

    const day = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10) - 1;
    let year = parseInt(dateParts[2], 10);

    if (year > 2500) {
      year -= 543;
    }

    const timeParts = timePart.split(":");
    const hours = parseInt(timeParts[0] || "0", 10);
    const minutes = parseInt(timeParts[1] || "0", 10);
    const seconds = parseInt(timeParts[2] || "0", 10);

    return new Date(year, month, day, hours, minutes, seconds);
  } catch (e) {
    return new Date(0); 
  }
};

/**
 * เรียกฟังก์ชันที่ยิง Sheets API (เช่น Sheets.Spreadsheets.Values.get /
 * batchGet) พร้อม retry แบบ exponential backoff เมื่อเจอ error
 * "Quota exceeded ... Read requests per minute" ซึ่งเป็น error ชั่วคราว
 * (rate-limit ต่อนาที ไม่ใช่โควต้ารายวัน) จะหายไปเองถ้ารอสักครู่แล้วลองใหม่
 * ไม่เปลี่ยน logic เดิม แค่ห่อ call เดิมด้วยการ retry ก่อนจะ throw จริงๆ
 * @param {Function} fn - ฟังก์ชันที่เรียก Sheets API แล้ว return ผลลัพธ์
 * @param {number} [maxRetries=4] - จำนวนครั้งที่ retry สูงสุด
 * @returns ผลลัพธ์จาก fn()
 */
function callSheetsApiWithBackoff(fn, maxRetries = 4) {
  let attempt = 0;
  while (true) {
    try {
      return fn();
    } catch (e) {
      const message = String((e && e.message) || e);
      const isQuotaError =
        message.indexOf("Quota exceeded") !== -1 ||
        message.indexOf("Read requests per minute") !== -1 ||
        message.indexOf("Rate Limit Exceeded") !== -1;

      if (!isQuotaError || attempt >= maxRetries) {
        throw e;
      }

      // Exponential backoff + jitter เล็กน้อย: 1s, 2s, 4s, 8s (+0-300ms)
      const delayMs = Math.pow(2, attempt) * 1000 + Math.floor(Math.random() * 300);
      Utilities.sleep(delayMs);
      attempt++;
    }
  }
}

// CacheService จำกัดขนาด value ต่อ key ไว้ที่ 100KB ถ้าข้อมูลที่จะแคชใหญ่กว่า
// นั้น cache.put() จะ throw "Argument too large: value" ทันที (เงียบๆ ถ้าอยู่ใน
// try/catch) ทำให้แคชไม่เคยถูกเขียนจริงเลยเมื่อข้อมูลเยอะ ฟังก์ชันคู่นี้จึง
// แบ่งข้อมูลเป็นชิ้นเล็กๆ (chunk) แล้วเก็บเป็นหลาย key แทน โดยไม่เปลี่ยน
// พฤติกรรมภายนอกของ caller (ยังคง get คืนค่าเดิมทุกประการ)
const CACHE_CHUNK_SIZE = 90000; // เผื่อ margin จากลิมิต 100KB ของ CacheService

/**
 * แคช string ขนาดใหญ่โดยแบ่งเป็นหลาย key ย่อยถ้าจำเป็น
 * @param {GoogleAppsScript.Cache.Cache} cache
 * @param {string} key
 * @param {string} value - ข้อมูลที่ต้องการแคช (ต้องเป็น string แล้ว เช่นจาก JSON.stringify)
 * @param {number} ttlSeconds
 */
function putCacheChunked(cache, key, value, ttlSeconds) {
  const totalChunks = Math.max(1, Math.ceil(value.length / CACHE_CHUNK_SIZE));
  const values = {};

  for (let i = 0; i < totalChunks; i++) {
    const chunkKey = key + "_chunk_" + i;
    values[chunkKey] = value.substring(
      i * CACHE_CHUNK_SIZE,
      (i + 1) * CACHE_CHUNK_SIZE,
    );
  }
  values[key + "_meta"] = JSON.stringify({ totalChunks: totalChunks });

  // putAll เขียนได้สูงสุดครั้งละ 100KB รวมทุก key เช่นกัน ถ้าจำนวน chunk
  // เยอะเกินไปให้ค่อยๆ put ทีละ key แทนเพื่อความปลอดภัย
  try {
    cache.putAll(values, ttlSeconds);
  } catch (e) {
    Object.keys(values).forEach((k) => {
      try {
        cache.put(k, values[k], ttlSeconds);
      } catch (innerErr) {
        // ถ้ายังใหญ่เกินไปแม้แบ่ง chunk แล้ว ปล่อยผ่าน (เหมือนพฤติกรรมเดิม
        // ที่ยอมให้ cache write ล้มเหลวได้โดยไม่กระทบผลลัพธ์หลัก)
      }
    });
  }
}

/**
 * อ่านค่าที่ถูกแคชไว้ด้วย putCacheChunked กลับมาเป็น string เดียวเหมือนเดิม
 * คืนค่า null ถ้าไม่พบหรือแคชไม่ครบ (เพื่อให้ caller ไปดึงข้อมูลใหม่เหมือน cache miss ปกติ)
 * @param {GoogleAppsScript.Cache.Cache} cache
 * @param {string} key
 * @returns {string|null}
 */
function getCacheChunked(cache, key) {
  const metaRaw = cache.get(key + "_meta");
  if (!metaRaw) return null;

  let totalChunks;
  try {
    totalChunks = JSON.parse(metaRaw).totalChunks;
  } catch (e) {
    return null;
  }

  const chunkKeys = [];
  for (let i = 0; i < totalChunks; i++) {
    chunkKeys.push(key + "_chunk_" + i);
  }

  const chunks = cache.getAll(chunkKeys);
  let result = "";
  for (let i = 0; i < totalChunks; i++) {
    const chunk = chunks[key + "_chunk_" + i];
    if (chunk === undefined || chunk === null) return null; // chunk ขาดหาย (หมดอายุไม่พร้อมกัน) ถือว่า cache miss
    result += chunk;
  }
  return result;
}