import * as fs from 'fs';
import * as path from 'path';
import {
  WorkBook,
  WorkSheet,
  ParsingOptions,
  WritingOptions,
  CellObject,
  CellStyle,
  Font,
  Fill,
  Border,
  Alignment,
  Range,
  CellAddress,
  PivotTableDefinition,
  PivotTableOptions,
  PivotDataField,
} from './types.js';
import {
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
  encode_col,
  decode_col,
  encode_row,
  decode_row,
} from './utils/address.js';
import {
  aoa_to_sheet,
  json_to_sheet,
  sheet_add_aoa,
  sheet_add_json,
  sheet_to_json,
  sheet_to_csv,
  sheet_to_txt,
  sheet_to_html,
  sheet_to_formulae,
  format_cell,
  set_cell_style,
  set_range_style,
} from './utils/sheet.js';
import { book_new, book_append_sheet, add_pivot_table } from './utils/book.js';
import { table_to_sheet, table_to_book } from './utils/table.js';
import { computePivotData, extractSheetData } from './pivot/compute.js';
import { writeXlsxSync, writeXlsxAsync, convertOutputType } from './xlsx/writer.js';
import { readXlsxSync, readXlsxAsync } from './xlsx/reader.js';
import { readCsv, writeCsv } from './csv/index.js';

export const version = '1.0.0';

function toUint8Array(data: any, type?: string): Uint8Array {
  if (data instanceof Uint8Array) {
    return data;
  }
  if (typeof Buffer !== 'undefined' && Buffer.isBuffer(data)) {
    return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (typeof data === 'string') {
    if (type === 'base64' || (!type && data.startsWith('UEsDB'))) {
      if (typeof Buffer !== 'undefined') {
        const buf = Buffer.from(data, 'base64');
        return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
      }
      const bin = atob(data);
      const u8 = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      return u8;
    }
    if (type === 'binary') {
      const u8 = new Uint8Array(data.length);
      for (let i = 0; i < data.length; i++) u8[i] = data.charCodeAt(i);
      return u8;
    }
    // String data utf-8
    const encoder = new TextEncoder();
    return encoder.encode(data);
  }
  if (Array.isArray(data)) {
    return new Uint8Array(data);
  }
  throw new Error(`Unsupported data type for read: ${typeof data}`);
}

function isZip(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export function read(data: any, opts?: ParsingOptions): WorkBook {
  const bytes = toUint8Array(data, opts?.type);

  if (isZip(bytes)) {
    return readXlsxSync(bytes, opts);
  }

  // Not a zip - try CSV/TSV or HTML
  const decoder = new TextDecoder('utf-8');
  const text = decoder.decode(bytes);

  if (text.trim().startsWith('<table') || text.trim().startsWith('<html')) {
    return table_to_book(text, opts);
  }

  return readCsv(text, opts);
}

export async function readAsync(data: any, opts?: ParsingOptions): Promise<WorkBook> {
  const bytes = toUint8Array(data, opts?.type);

  if (isZip(bytes)) {
    return await readXlsxAsync(bytes, opts);
  }

  const decoder = new TextDecoder('utf-8');
  const text = decoder.decode(bytes);

  if (text.trim().startsWith('<table') || text.trim().startsWith('<html')) {
    return table_to_book(text, opts);
  }

  return readCsv(text, opts);
}

export function readFile(filename: string, opts?: ParsingOptions): WorkBook {
  const buffer = fs.readFileSync(filename);
  return read(buffer, opts);
}

export async function readFileAsync(filename: string, opts?: ParsingOptions): Promise<WorkBook> {
  const buffer = await fs.promises.readFile(filename);
  return await readAsync(buffer, opts);
}

export function write(wb: WorkBook, opts?: WritingOptions): any {
  const bookType = opts?.bookType || 'xlsx';

  if (bookType === 'csv' || bookType === 'txt') {
    const csvStr = writeCsv(wb, opts);
    if (opts?.type === 'buffer' && typeof Buffer !== 'undefined') {
      return Buffer.from(csvStr, 'utf-8');
    }
    if (opts?.type === 'array') {
      return new TextEncoder().encode(csvStr);
    }
    return csvStr;
  }

  if (bookType === 'html') {
    const firstSheet = wb.SheetNames[0] || 'Sheet1';
    const htmlStr = sheet_to_html(wb.Sheets[firstSheet] || {});
    if (opts?.type === 'buffer' && typeof Buffer !== 'undefined') {
      return Buffer.from(htmlStr, 'utf-8');
    }
    if (opts?.type === 'array') {
      return new TextEncoder().encode(htmlStr);
    }
    return htmlStr;
  }

  if (bookType === 'json') {
    const result: Record<string, any[]> = {};
    for (const sName of wb.SheetNames) {
      result[sName] = sheet_to_json(wb.Sheets[sName] || {});
    }
    const jsonStr = JSON.stringify(result, null, 2);
    if (opts?.type === 'buffer' && typeof Buffer !== 'undefined') {
      return Buffer.from(jsonStr, 'utf-8');
    }
    return jsonStr;
  }

  // Default XLSX
  const bytes = writeXlsxSync(wb, opts);
  return convertOutputType(bytes, opts?.type);
}

export async function writeAsync(wb: WorkBook, opts?: WritingOptions): Promise<any> {
  const bookType = opts?.bookType || 'xlsx';

  if (bookType !== 'xlsx') {
    return write(wb, opts);
  }

  const bytes = await writeXlsxAsync(wb, opts);
  return convertOutputType(bytes, opts?.type);
}

export function writeFile(wb: WorkBook, filename: string, opts?: WritingOptions): void {
  const ext = path.extname(filename).toLowerCase();
  let bookType = opts?.bookType;
  if (!bookType) {
    if (ext === '.csv') bookType = 'csv';
    else if (ext === '.txt' || ext === '.tsv') bookType = 'txt';
    else if (ext === '.html' || ext === '.htm') bookType = 'html';
    else if (ext === '.json') bookType = 'json';
    else bookType = 'xlsx';
  }

  const data = write(wb, { ...opts, bookType, type: 'buffer' });
  fs.writeFileSync(filename, data);
}

export const writeFileXLSX = writeFile;

export async function writeFileAsync(
  wb: WorkBook,
  filename: string,
  opts?: WritingOptions
): Promise<void> {
  const ext = path.extname(filename).toLowerCase();
  let bookType = opts?.bookType;
  if (!bookType) {
    if (ext === '.csv') bookType = 'csv';
    else if (ext === '.txt' || ext === '.tsv') bookType = 'txt';
    else if (ext === '.html' || ext === '.htm') bookType = 'html';
    else if (ext === '.json') bookType = 'json';
    else bookType = 'xlsx';
  }

  const data = await writeAsync(wb, { ...opts, bookType, type: 'buffer' });
  await fs.promises.writeFile(filename, data);
}

export const utils = {
  // Address helpers
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
  encode_col,
  decode_col,
  encode_row,
  decode_row,

  // Sheet creation & conversion
  aoa_to_sheet,
  json_to_sheet,
  table_to_sheet,
  table_to_book,
  sheet_to_json,
  sheet_to_csv,
  sheet_to_txt,
  sheet_to_html,
  sheet_to_formulae,
  sheet_to_row_object_array: sheet_to_json,
  sheet_add_aoa,
  sheet_add_json,

  // Book helpers
  book_new,
  book_append_sheet,
  format_cell,

  // STYLING EXTENSIONS
  set_cell_style,
  set_range_style,

  // PIVOT TABLE EXTENSIONS
  add_pivot_table,
  compute_pivot_table: (data: Array<Record<string, any>>, options: any) =>
    computePivotData(data, options),
};

const XLSX = {
  version,
  read,
  readAsync,
  readFile,
  readFileAsync,
  write,
  writeAsync,
  writeFile,
  writeFileXLSX,
  writeFileAsync,
  utils,
};

export default XLSX;
export * from './types.js';
