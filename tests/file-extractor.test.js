/**
 * Unit tests for wecom/file-extractor.js
 *
 * Tests XLSX extraction, CSV encoding detection, file block formatting,
 * and the entry-point dispatch logic.
 */

import { describe, it, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { writeFileSync, mkdirSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

import {
  extractXlsxToText,
  extractCsvToText,
  formatFileBlock,
  tryExtractFileContent,
} from "../wecom/file-extractor.js";

// ── Helpers ────────────────────────────────────────────────────────────

const TEST_DIR = join(tmpdir(), `file-extractor-test-${Date.now()}`);

function setupTestDir() {
  mkdirSync(TEST_DIR, { recursive: true });
}

function cleanupTestDir() {
  try {
    rmSync(TEST_DIR, { recursive: true, force: true });
  } catch {
    // ignore
  }
}

function writeTempFile(name, content) {
  const filePath = join(TEST_DIR, name);
  writeFileSync(filePath, content);
  return filePath;
}

// ── formatFileBlock ────────────────────────────────────────────────────

describe("formatFileBlock", () => {
  it("produces gateway-compatible <file> block", () => {
    const result = formatFileBlock("test.csv", "text/csv", "a,b,c\n1,2,3");
    assert.equal(result, '<file name="test.csv" mime="text/csv">\na,b,c\n1,2,3\n</file>');
  });
});

// ── CSV extraction ─────────────────────────────────────────────────────

describe("extractCsvToText", () => {
  it("extracts UTF-8 CSV content", async () => {
    const buf = Buffer.from("name,value\nAlice,100\nBob,200", "utf8");
    const result = await extractCsvToText(buf, "data.csv");
    assert.ok(result);
    assert.ok(result.includes("Alice"));
    assert.ok(result.includes("Bob"));
  });

  it("handles UTF-8 BOM", async () => {
    const bom = Buffer.from([0xef, 0xbb, 0xbf]);
    const content = Buffer.from("name,value\ntest,123", "utf8");
    const buf = Buffer.concat([bom, content]);
    const result = await extractCsvToText(buf, "bom.csv");
    assert.ok(result);
    assert.ok(result.includes("name,value"));
    assert.ok(!result.startsWith("\ufeff")); // BOM should be stripped
  });

  it("returns null for empty buffer", async () => {
    const result = await extractCsvToText(Buffer.alloc(0), "empty.csv");
    assert.equal(result, null);
  });

  it("returns null for whitespace-only content", async () => {
    const result = await extractCsvToText(Buffer.from("   \n  \n  ", "utf8"), "blank.csv");
    assert.equal(result, null);
  });

  it("truncates content exceeding maxChars", async () => {
    const line = "a".repeat(100) + "\n";
    const buf = Buffer.from(line.repeat(50), "utf8");
    const result = await extractCsvToText(buf, "big.csv", { maxChars: 200 });
    assert.ok(result);
    assert.ok(result.length <= 220); // 200 + truncation marker
    assert.ok(result.includes("[truncated]"));
  });

  it("handles GBK-encoded CSV when iconv-lite is available", async () => {
    // Try loading iconv-lite; skip test if not installed
    let iconv;
    try {
      iconv = await import("iconv-lite");
    } catch {
      return; // skip — iconv-lite not installed
    }

    const gbkBuf = iconv.encode("姓名,数值\n张三,100\n李四,200", "gbk");
    const result = await extractCsvToText(gbkBuf, "gbk.csv");
    assert.ok(result);
    assert.ok(result.includes("张三"), "Should decode GBK Chinese characters");
    assert.ok(result.includes("李四"));
  });
});

// ── XLSX extraction ────────────────────────────────────────────────────

describe("extractXlsxToText", () => {
  let XLSX;

  beforeEach(async () => {
    try {
      XLSX = await import("xlsx");
    } catch {
      XLSX = null;
    }
  });

  it("returns a string (or null) for any buffer — never throws", async () => {
    // xlsx is lenient and may parse arbitrary bytes as a single-cell sheet.
    // The key contract is: no unhandled exception.
    const corrupt = Buffer.from([0xde, 0xad, 0xbe, 0xef, 0x00, 0x01, 0x02, 0x03]);
    const result = await extractXlsxToText(corrupt, "corrupt.xlsx");
    assert.ok(result === null || typeof result === "string");
  });

  it("extracts single-sheet XLSX content", async () => {
    if (!XLSX) return; // skip — xlsx not installed

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["Name", "Score"],
      ["Alice", 95],
      ["Bob", 87],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "Results");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const result = await extractXlsxToText(buf, "scores.xlsx");
    assert.ok(result);
    assert.ok(result.includes("Alice"));
    assert.ok(result.includes("95"));
    assert.ok(!result.includes("--- Sheet:")); // single sheet, no header
  });

  it("extracts multi-sheet XLSX with sheet headers", async () => {
    if (!XLSX) return;

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([["A", "B"], [1, 2]]);
    const ws2 = XLSX.utils.aoa_to_sheet([["C", "D"], [3, 4]]);
    XLSX.utils.book_append_sheet(wb, ws1, "Sheet1");
    XLSX.utils.book_append_sheet(wb, ws2, "Sheet2");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const result = await extractXlsxToText(buf, "multi.xlsx");
    assert.ok(result);
    assert.ok(result.includes("--- Sheet: Sheet1 ---"));
    assert.ok(result.includes("--- Sheet: Sheet2 ---"));
  });

  it("handles CJK content in XLSX", async () => {
    if (!XLSX) return;

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ["姓名", "分数"],
      ["张三", 100],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, "成绩");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const result = await extractXlsxToText(buf, "cjk.xlsx");
    assert.ok(result);
    assert.ok(result.includes("张三"));
    assert.ok(result.includes("100"));
  });

  it("truncates large XLSX content", async () => {
    if (!XLSX) return;

    const rows = [["Col1", "Col2"]];
    for (let i = 0; i < 500; i++) {
      rows.push([`Row${i}`, "x".repeat(100)]);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "Big");
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

    const result = await extractXlsxToText(buf, "big.xlsx", { maxChars: 500 });
    assert.ok(result);
    assert.ok(result.length <= 520);
    assert.ok(result.includes("[truncated]"));
  });
});

// ── tryExtractFileContent (entry point) ────────────────────────────────

describe("tryExtractFileContent", () => {
  beforeEach(() => {
    setupTestDir();
  });

  it("returns null for non-extractable extensions (.pdf)", async () => {
    const filePath = writeTempFile("doc.pdf", "dummy pdf content");
    const result = await tryExtractFileContent(filePath, "doc.pdf");
    assert.equal(result, null);
  });

  it("returns null for image files", async () => {
    const filePath = writeTempFile("photo.jpg", "not really an image");
    const result = await tryExtractFileContent(filePath, "photo.jpg");
    assert.equal(result, null);
  });

  it("returns null for .docx files (not yet supported)", async () => {
    const filePath = writeTempFile("doc.docx", "fake docx");
    const result = await tryExtractFileContent(filePath, "doc.docx");
    assert.equal(result, null);
  });

  it("extracts CSV file content", async () => {
    const filePath = writeTempFile("data.csv", "a,b,c\n1,2,3\n4,5,6");
    const result = await tryExtractFileContent(filePath, "data.csv");
    assert.ok(result);
    assert.ok(result.fileBlock.includes('<file name="data.csv"'));
    assert.ok(result.fileBlock.includes("text/csv"));
    assert.ok(result.text.includes("1,2,3"));
  });

  it("extracts XLSX file content when xlsx library is available", async () => {
    let XLSX;
    try {
      XLSX = await import("xlsx");
    } catch {
      return; // skip — xlsx not installed
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet([["Key", "Val"], ["hello", 42]]),
      "Data",
    );
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    const filePath = writeTempFile("test.xlsx", buf);

    const result = await tryExtractFileContent(filePath, "test.xlsx");
    assert.ok(result);
    assert.ok(result.fileBlock.includes("hello"));
    assert.ok(result.fileBlock.includes("42"));
    assert.ok(result.fileBlock.includes("spreadsheetml"));
  });

  it("returns null for empty file", async () => {
    const filePath = writeTempFile("empty.csv", "");
    const result = await tryExtractFileContent(filePath, "empty.csv");
    assert.equal(result, null);
  });

  it("returns null for nonexistent file path", async () => {
    const result = await tryExtractFileContent("/nonexistent/path/file.csv", "file.csv");
    assert.equal(result, null);
  });

  it("handles .xls extension", async () => {
    let XLSX;
    try {
      XLSX = await import("xlsx");
    } catch {
      return;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.aoa_to_sheet([["A"], [1]]),
      "S1",
    );
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xls" });
    const filePath = writeTempFile("old.xls", buf);

    const result = await tryExtractFileContent(filePath, "old.xls");
    assert.ok(result);
    assert.ok(result.fileBlock.includes("vnd.ms-excel"));
  });

  // Cleanup
  it("cleanup test directory", () => {
    cleanupTestDir();
  });
});
