/**
 * File content extraction for XLSX and CSV files.
 *
 * Extracts text content from binary spreadsheet and CSV files BEFORE
 * dispatching to the gateway, bypassing the gateway's `isBinaryMediaMime()`
 * limitation that blocks `application/vnd.*` MIME types.
 *
 * Dependencies are optional — the plugin falls back to the existing
 * gateway pipeline when extraction libraries are unavailable.
 */

import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import { logger } from "../logger.js";

const DEFAULT_MAX_CHARS = 200_000;

// ── Dynamic imports with graceful fallback ──────────────────────────────

let xlsxModule = undefined;
let iconvModule = undefined;

async function loadXlsx() {
  if (xlsxModule !== undefined) return xlsxModule;
  try {
    xlsxModule = await import("xlsx");
    return xlsxModule;
  } catch {
    xlsxModule = null;
    return null;
  }
}

async function loadIconvLite() {
  if (iconvModule !== undefined) return iconvModule;
  try {
    iconvModule = await import("iconv-lite");
    return iconvModule;
  } catch {
    iconvModule = null;
    return null;
  }
}

// ── XLSX extraction ─────────────────────────────────────────────────────

/**
 * Parse an XLSX buffer into CSV text per sheet.
 * Returns null if xlsx library is unavailable or the file is corrupt.
 */
export async function extractXlsxToText(buffer, fileName, { maxChars = DEFAULT_MAX_CHARS } = {}) {
  const XLSX = await loadXlsx();
  if (!XLSX) {
    logger.debug("[file-extractor] xlsx library not available, skipping extraction");
    return null;
  }

  try {
    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sheets = workbook.SheetNames;
    if (!sheets || sheets.length === 0) return null;

    const parts = [];
    let totalChars = 0;

    for (const name of sheets) {
      const sheet = workbook.Sheets[name];
      const csv = XLSX.utils.sheet_to_csv(sheet);
      if (!csv.trim()) continue;

      const header = sheets.length > 1 ? `--- Sheet: ${name} ---\n` : "";
      const section = header + csv;

      if (totalChars + section.length > maxChars) {
        const remaining = maxChars - totalChars;
        if (remaining > 0) {
          parts.push(section.substring(0, remaining) + "\n...[truncated]");
        }
        break;
      }

      parts.push(section);
      totalChars += section.length;
    }

    const text = parts.join("\n\n");
    return text || null;
  } catch (err) {
    logger.warn("[file-extractor] XLSX parse failed", { fileName, error: err.message });
    return null;
  }
}

// ── CSV extraction ──────────────────────────────────────────────────────

/**
 * Detect whether a buffer is likely GBK/GB18030 encoded.
 * Heuristic: check for high-byte pairs common in GBK but invalid in UTF-8.
 */
function looksLikeGbk(buffer) {
  let invalidUtf8Sequences = 0;
  let gbkLikePairs = 0;
  const len = Math.min(buffer.length, 4096);

  for (let i = 0; i < len; i++) {
    const b = buffer[i];
    if (b < 0x80) continue;

    // Check if it looks like a valid UTF-8 multi-byte sequence
    if (b >= 0xc0 && b <= 0xdf && i + 1 < len) {
      const b2 = buffer[i + 1];
      if (b2 >= 0x80 && b2 <= 0xbf) {
        i += 1;
        continue;
      }
    }
    if (b >= 0xe0 && b <= 0xef && i + 2 < len) {
      const b2 = buffer[i + 1];
      const b3 = buffer[i + 2];
      if (b2 >= 0x80 && b2 <= 0xbf && b3 >= 0x80 && b3 <= 0xbf) {
        i += 2;
        continue;
      }
    }

    // If we get here with high byte, it's likely GBK
    if (b >= 0x81 && b <= 0xfe && i + 1 < len) {
      const b2 = buffer[i + 1];
      if (b2 >= 0x40 && b2 <= 0xfe) {
        gbkLikePairs++;
        i += 1;
        continue;
      }
    }

    invalidUtf8Sequences++;
  }

  return gbkLikePairs > 0 && invalidUtf8Sequences + gbkLikePairs > 2;
}

/**
 * Decode a CSV buffer to text, handling GBK/GB18030 encoding.
 * Falls back to UTF-8 → iconv-lite GB18030 → Latin-1.
 */
export async function extractCsvToText(buffer, fileName, { maxChars = DEFAULT_MAX_CHARS } = {}) {
  let text;

  // Strip BOM if present
  const hasBom = buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
  const cleanBuffer = hasBom ? buffer.subarray(3) : buffer;

  if (looksLikeGbk(cleanBuffer)) {
    const iconv = await loadIconvLite();
    if (iconv) {
      try {
        text = iconv.decode(cleanBuffer, "gb18030");
        logger.debug("[file-extractor] CSV decoded as GB18030", { fileName });
      } catch {
        text = cleanBuffer.toString("utf8");
      }
    } else {
      // No iconv-lite: try utf-8, fall back to latin1
      text = cleanBuffer.toString("utf8");
      if (text.includes("\ufffd")) {
        text = cleanBuffer.toString("latin1");
        logger.debug("[file-extractor] CSV fallback to latin1 (no iconv-lite)", { fileName });
      }
    }
  } else {
    text = cleanBuffer.toString("utf8");
  }

  if (!text || !text.trim()) return null;

  if (text.length > maxChars) {
    return text.substring(0, maxChars) + "\n...[truncated]";
  }

  return text;
}

// ── File block formatting ───────────────────────────────────────────────

/**
 * Format extracted text as a `<file>` block matching the gateway's
 * `extractFileBlocks()` output format.
 */
export function formatFileBlock(fileName, mimeType, text) {
  return `<file name="${fileName}" mime="${mimeType}">\n${text}\n</file>`;
}

// ── Entry point ─────────────────────────────────────────────────────────

const EXTRACTABLE_EXTENSIONS = new Set([".xlsx", ".xls", ".csv"]);

/**
 * Attempt to extract text content from a local file.
 *
 * @param {string} localFilePath - Absolute path to the downloaded file
 * @param {string} fileName - Original file name (for extension detection)
 * @param {object} [options]
 * @param {number} [options.maxChars=200000] - Maximum characters to extract
 * @returns {Promise<{ fileBlock: string, text: string } | null>}
 *   Returns null when extraction is not applicable or fails.
 */
export async function tryExtractFileContent(localFilePath, fileName, { maxChars = DEFAULT_MAX_CHARS } = {}) {
  const ext = extname(fileName || localFilePath).toLowerCase();

  if (!EXTRACTABLE_EXTENSIONS.has(ext)) {
    return null;
  }

  let buffer;
  try {
    buffer = await readFile(localFilePath);
  } catch (err) {
    logger.warn("[file-extractor] failed to read file", { path: localFilePath, error: err.message });
    return null;
  }

  if (buffer.length === 0) return null;

  let text = null;
  let mimeType = "application/octet-stream";

  if (ext === ".xlsx" || ext === ".xls") {
    text = await extractXlsxToText(buffer, fileName, { maxChars });
    mimeType = ext === ".xlsx"
      ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      : "application/vnd.ms-excel";
  } else if (ext === ".csv") {
    text = await extractCsvToText(buffer, fileName, { maxChars });
    mimeType = "text/csv";
  }

  if (!text) return null;

  const fileBlock = formatFileBlock(fileName || localFilePath, mimeType, text);
  return { fileBlock, text };
}
