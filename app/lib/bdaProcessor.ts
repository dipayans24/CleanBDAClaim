// lib/bdaProcessor.ts
//
// This file is a line-by-line port of the original Python/pandas script into
// TypeScript, so it can run inside a Vercel serverless function (Node.js
// runtime). Every function below has a comment tying it back to the exact
// step it replaces in the Python code.

import * as XLSX from "xlsx"; // reads/writes both .csv and .xlsx files
import dayjs from "dayjs"; // small date library
import customParseFormat from "dayjs/plugin/customParseFormat"; // lets dayjs parse a custom string format

dayjs.extend(customParseFormat); // turn on the plugin so dayjs(str, format, strict) works

// A "row" is just a plain object keyed by column name, e.g. { phoneNumber: 123, ... }
export type BdaRow = Record<string, any>;

// ---------------------------------------------------------------------------
// Helper: pythonInt
// Mirrors the Python `convert_to_int` function:
//   def convert_to_int(x):
//       try: return int(x)
//       except: return x
// Python's int() has specific rules we need to copy exactly:
//   - int(float) truncates toward zero (e.g. int(-3.7) == -3, not -4)
//   - int("123") works, but int("123.0") FAILS (raises ValueError)
//   - int(None) / int(NaN) FAILS
// On failure, the original value is returned unchanged (that's the `except` branch).
// ---------------------------------------------------------------------------
export function pythonInt(x: unknown): unknown {
  if (typeof x === "number") {
    // NaN/Infinity can't be converted in Python either -> falls into "except", return unchanged
    if (!Number.isFinite(x)) return x;
    return Math.trunc(x); // truncate toward zero, same as Python's int(float)
  }
  if (typeof x === "string") {
    const trimmed = x.trim(); // Python's int() tolerates surrounding whitespace
    // Only a plain integer literal (optional +/-, digits only) is accepted -
    // "123.0", "12a", "" etc. all fail, just like Python's int(str)
    if (/^[+-]?\d+$/.test(trimmed)) {
      return parseInt(trimmed, 10);
    }
    return x; // conversion failed -> return the original string, unchanged
  }
  // Anything else (null, undefined, boolean, object) can't be converted -> return unchanged
  return x;
}

// ---------------------------------------------------------------------------
// Helper: pyStr
// Mirrors Python's str(x) specifically for the "missing value" case, where
// pandas stores missing data as float('nan') and str(nan) == "nan".
// Used only for the alternatePhoneNumbers length check below.
// ---------------------------------------------------------------------------
function pyStr(x: unknown): string {
  if (x === null || x === undefined) return "nan"; // matches str(NaN) in pandas
  if (typeof x === "number" && Number.isNaN(x)) return "nan";
  return String(x);
}

// ---------------------------------------------------------------------------
// Helper: parseStrictDateTime
// Mirrors:
//   pd.to_datetime(col, format="%d/%m/%Y %H:%M", exact=True, dayfirst=True, yearfirst=False)
// "exact=True" means the ENTIRE string must match the format, not just a
// prefix - dayjs's strict-parsing mode (the 3rd argument = true) does the same.
// If a value doesn't match, pandas raises an error rather than silently
// producing a null, so we throw here too (with a helpful row/column reference)
// instead of quietly corrupting the output.
// ---------------------------------------------------------------------------
function parseStrictDateTime(raw: unknown, rowIndex: number, columnName: string): Date {
  const text = raw === null || raw === undefined ? "" : String(raw).trim();
  const parsed = dayjs(text, "DD/MM/YYYY HH:mm", true); // true = strict mode
  if (!parsed.isValid()) {
    // rowIndex + 2 approximates the row number a person would see in Excel
    // (row 1 is the header, and arrays are 0-indexed)
    throw new Error(
      `Row ${rowIndex + 2}: column "${columnName}" value "${text}" does not match the expected format DD/MM/YYYY HH:mm`
    );
  }
  return parsed.toDate();
}

// ---------------------------------------------------------------------------
// Helper: toNumberIfNumeric
// CSV files have no cell types - every value comes back as a plain string.
// pandas' read_csv/read_excel auto-detects numeric columns and gives you real
// int/float values, but our CSV parser does not, so the script's `== 0` and
// arithmetic checks (phoneNumber, paymentAmount, refundAmount, closing_amt,
// rest_collection) would silently never match. This coerces any string that
// looks like a plain number into an actual JS number; anything else (blank,
// text) is left untouched.
// ---------------------------------------------------------------------------
function toNumberIfNumeric(x: unknown): unknown {
  if (typeof x === "number" || x === null || x === undefined) return x;
  if (typeof x === "string") {
    const trimmed = x.trim();
    if (trimmed !== "" && /^[+-]?\d+(\.\d+)?$/.test(trimmed)) {
      return Number(trimmed);
    }
  }
  return x;
}

// ---------------------------------------------------------------------------
// readInputFile
// Mirrors:
//   if extn == "csv": BDA = pd.read_csv(BDAFilePath, sep=",")
//   else: BDA = pd.read_excel(BDAFilePath)
// The xlsx library auto-detects CSV vs. binary Excel content, so one call
// handles both cases - we don't need to branch on the file extension here.
// ---------------------------------------------------------------------------
export function readInputFile(buffer: Buffer, _originalName: string): BdaRow[] {
  // Parse the raw bytes into a workbook object (works for .csv, .xlsx, .xls)
  const workbook = XLSX.read(buffer, { type: "buffer", raw: true, cellDates: false });

  // Use the first sheet, same as pandas defaults to the first sheet of an Excel file
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("The uploaded file does not contain any sheets/data.");
  }
  const sheet = workbook.Sheets[firstSheetName];

  // Convert the sheet into an array of { columnName: value } row objects.
  // defval: null makes genuinely empty cells become `null` instead of being
  // dropped from the row entirely, which keeps every row's keys consistent.
  const rows: BdaRow[] = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  return rows;
}

// ---------------------------------------------------------------------------
// processBda
// This is the main pipeline - every block below maps to one paragraph of the
// original Python script, in the same order.
// ---------------------------------------------------------------------------
export interface ProcessBdaResult {
  rows: BdaRow[];
  columnOrder: string[];
  emptyPhoneCount: number;
}

export function processBda(inputRows: BdaRow[]): ProcessBdaResult {
  if (!inputRows || inputRows.length === 0) {
    throw new Error("The uploaded file has no data rows.");
  }

  // Capture the source column order (JS objects preserve string-key insertion
  // order, so Object.keys() gives us the same left-to-right order as the file)
  const originalColumns = Object.keys(inputRows[0]);

  // Fail fast with a clear message if a column the script depends on is missing,
  // rather than crashing later with a confusing "undefined" error
  const requiredColumns = [
    "createdAt",
    "updatedAt",
    "date_time_of_closing",
    "original_phone_number",
    "phoneNumber",
    "paymentStatus",
    "paymentAmount",
    "refundAmount",
    "closing_amt",
    "rest_collection",
    "alternateEmails",
    "alternatePhoneNumbers",
    "paymentId",
    "paymentproof",
    "remark",
  ];
  for (const col of requiredColumns) {
    if (!originalColumns.includes(col)) {
      throw new Error(`Missing required column "${col}" in the uploaded file.`);
    }
  }

  // Work on copies of every row so the caller's original array is never mutated
  let BDA: BdaRow[] = inputRows.map((row) => ({ ...row }));

  // ---- make sure the columns the script does arithmetic/equality checks on are real numbers ----
  // (pandas would have already inferred these as int/float columns on read)
  const numericColumns = ["phoneNumber", "paymentAmount", "refundAmount", "closing_amt", "rest_collection"];
  BDA.forEach((row) => {
    for (const col of numericColumns) {
      row[col] = toNumberIfNumeric(row[col]);
    }
  });

  // ---- BDA["createdAt"] / ["updatedAt"] / ["date_time_of_closing"] = pd.to_datetime(...) ----
  BDA = BDA.map((row, i) => {
    row.createdAt = parseStrictDateTime(row.createdAt, i, "createdAt");
    row.updatedAt = parseStrictDateTime(row.updatedAt, i, "updatedAt");
    row.date_time_of_closing = parseStrictDateTime(row.date_time_of_closing, i, "date_time_of_closing");
    return row;
  });

  // ---- BDA["original_phone_number"] = BDA["original_phone_number"].map(convert_to_int) ----
  BDA.forEach((row) => {
    row.original_phone_number = pythonInt(row.original_phone_number);
  });

  // ---- BDA["phoneNumber"] = original_phone_number when phoneNumber == 0 ----
  BDA.forEach((row) => {
    if (row.phoneNumber === 0) {
      row.phoneNumber = row.original_phone_number;
    }
  });

  // ---- fill any phoneNumber values that are STILL 0 with 101, 102, 103, ... ----
  // (this only happens when original_phone_number was itself 0/missing)
  let counter = 101; // matches Python's range(101, 101 + len(empty_phone))
  let emptyPhoneCount = 0;
  BDA.forEach((row) => {
    if (row.phoneNumber === 0) {
      emptyPhoneCount += 1;
      row.phoneNumber = counter;
      counter += 1;
    }
  });
  // eslint-disable-next-line no-console -- mirrors the original script's print() statement
  console.log(`Total number of phonenumber with '0'- ${emptyPhoneCount}.`);

  // ---- partial_refunded = BDA[BDA["paymentStatus"].str.contains('partial_refunded')] ----
  const partialRefunded: BdaRow[] = BDA.filter(
    (row) => typeof row.paymentStatus === "string" && row.paymentStatus.includes("partial_refunded")
  ).map((row) => ({ ...row })); // copy, so edits below don't leak back into BDA

  // ---- partial_refunded["FinalPayment"] = paymentAmount - refundAmount ----
  partialRefunded.forEach((row) => {
    const paymentAmount = typeof row.paymentAmount === "number" ? row.paymentAmount : 0;
    const refundAmount = typeof row.refundAmount === "number" ? row.refundAmount : 0;
    row.FinalPayment = paymentAmount - refundAmount;
  });

  // ---- BDA = BDA[~BDA["paymentStatus"].str.contains('refund')] ----
  // (removes both full refunds AND the partial refunds we already copied out above)
  BDA = BDA.filter((row) => !(typeof row.paymentStatus === "string" && row.paymentStatus.includes("refund")));

  // ---- BDA.insert(loc=paymentAmountLoc, column="FinalPayment", value=max(closing_amt, rest_collection) if paymentAmount==0 else paymentAmount) ----
  BDA.forEach((row) => {
    if (row.paymentAmount === 0) {
      const closingAmt = typeof row.closing_amt === "number" ? row.closing_amt : 0;
      const restCollection = typeof row.rest_collection === "number" ? row.rest_collection : 0;
      row.FinalPayment = Math.max(closingAmt, restCollection);
    } else {
      row.FinalPayment = row.paymentAmount;
    }
  });

  // Work out where "FinalPayment" should sit in the column order: directly
  // before "paymentAmount", same as BDA.insert(loc=FinalPaymentLoc, ...)
  const paymentAmountIndex = originalColumns.indexOf("paymentAmount");
  const columnsWithFinalPayment = [
    ...originalColumns.slice(0, paymentAmountIndex),
    "FinalPayment",
    ...originalColumns.slice(paymentAmountIndex),
  ];

  // ---- final_BDA = pd.concat([BDA, partial_refunded], axis="rows", ignore_index=True) ----
  const finalBDA: BdaRow[] = [...BDA, ...partialRefunded];

  // ---- final_BDA["alternateEmails"] = ... keep if len(x) < 71, else NA ----
  finalBDA.forEach((row) => {
    const text = row.alternateEmails === null || row.alternateEmails === undefined ? "" : String(row.alternateEmails);
    if (text.length >= 71) {
      row.alternateEmails = null; // equivalent to pandas' pd.NA
    }
  });

  // ---- final_BDA["alternatePhoneNumbers"] = ... keep if len(str(x)) < 71, else NA ----
  finalBDA.forEach((row) => {
    const text = pyStr(row.alternatePhoneNumbers); // str() of the raw value, "nan" if missing
    if (text.length >= 71) {
      row.alternatePhoneNumbers = null;
    }
  });

  // ---- final_BDA["paymentId"] = final_BDA["paymentId"].map(convert_to_int) ----
  finalBDA.forEach((row) => {
    row.paymentId = pythonInt(row.paymentId);
  });

  // ---- final_BDA.drop(columns=["paymentproof", "remark"], inplace=True) ----
  finalBDA.forEach((row) => {
    delete row.paymentproof;
    delete row.remark;
  });
  const finalColumnOrder = columnsWithFinalPayment.filter((col) => col !== "paymentproof" && col !== "remark");

  return { rows: finalBDA, columnOrder: finalColumnOrder, emptyPhoneCount };
}

// ---------------------------------------------------------------------------
// writeOutputWorkbook
// Mirrors:
//   final_BDA.to_excel(rf"{filedir}\{fileName}_cleaned.xlsx", index=False)
// Builds an in-memory .xlsx file (as a Buffer) ready to send back to the browser.
// ---------------------------------------------------------------------------
export function writeOutputWorkbook(rows: BdaRow[], columnOrder: string[]): Buffer {
  // header: columnOrder forces the exact column order/set we computed above,
  // instead of letting the library guess it from object key order
  const worksheet = XLSX.utils.json_to_sheet(rows, { header: columnOrder });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  const outputBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
  return outputBuffer;
}
