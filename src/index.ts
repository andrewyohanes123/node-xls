import * as fs from 'fs';
import * as path from 'path';
import * as SSF from 'ssf';
import { WorkBook, ParsingOptions, WritingOptions, BookType } from './types.js';
import {
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
  encode_col,
  decode_col,
  encode_row,
  decode_row,
  split_cell,
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
  sheet_new,
  sheet_get_cell,
  cell_set_number_format,
  cell_set_hyperlink,
  cell_set_internal_link,
  cell_add_comment,
  sheet_set_array_formula,
} from './utils/sheet.js';
import { book_new, book_append_sheet, book_set_sheet_visibility, add_pivot_table } from './utils/book.js';
import { table_to_sheet, table_to_book } from './utils/table.js';
import { computePivotData } from './pivot/compute.js';
import { writeXlsxSync, writeXlsxAsync, convertOutputType } from './xlsx/writer.js';
import { readXlsxSync, readXlsxAsync } from './xlsx/reader.js';
import { readCsv, writeCsv, writeHtml } from './csv/index.js';
import { dateToSerial, serialToDate } from './utils/dates.js';

export const version = '0.1.0';
export { SSF };

function toUint8Array(data: any, type?: string): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (typeof data === 'string') {
    if (type === 'base64' || (!type && data.startsWith('UEsDB'))) {
      const buf = Buffer.from(data, 'base64');
      return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    }
    if (type === 'binary') {
      const u8 = new Uint8Array(data.length);
      for (let i = 0; i < data.length; i++) u8[i] = data.charCodeAt(i) & 0xff;
      return u8;
    }
    return new TextEncoder().encode(data);
  }
  if (Array.isArray(data)) return new Uint8Array(data);
  throw new Error(`Unsupported data type for read: ${typeof data}`);
}

const isZip = (b: Uint8Array) => b.length >= 4 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;
const looksLikeHtml = (text: string) => /^\s*(<!doctype html|<html|<table|<body)/i.test(text);

function readText(bytes: Uint8Array, opts?: ParsingOptions): WorkBook {
  const text = new TextDecoder('utf-8').decode(bytes);
  return looksLikeHtml(text) ? table_to_book(text, opts as any) : readCsv(text, opts);
}

export function read(data: any, opts?: ParsingOptions): WorkBook {
  const bytes = toUint8Array(data, opts?.type);
  return isZip(bytes) ? readXlsxSync(bytes, opts) : readText(bytes, opts);
}

export async function readAsync(data: any, opts?: ParsingOptions): Promise<WorkBook> {
  const bytes = toUint8Array(data, opts?.type);
  return isZip(bytes) ? readXlsxAsync(bytes, opts) : readText(bytes, opts);
}

export function readFile(filename: string, opts?: ParsingOptions): WorkBook {
  return read(fs.readFileSync(filename), opts);
}

export async function readFileAsync(filename: string, opts?: ParsingOptions): Promise<WorkBook> {
  return readAsync(await fs.promises.readFile(filename), opts);
}

function encodeText(text: string, type?: string): any {
  if (type === 'buffer' || type === 'file') return Buffer.from(text, 'utf-8');
  if (type === 'array') return new TextEncoder().encode(text);
  if (type === 'base64') return Buffer.from(text, 'utf-8').toString('base64');
  return text;
}

export function write(wb: WorkBook, opts?: WritingOptions): any {
  const bookType = opts?.bookType || 'xlsx';
  switch (bookType) {
    case 'csv':
    case 'txt':
      return encodeText(writeCsv(wb, opts), opts?.type);
    case 'html':
      return encodeText(writeHtml(wb, opts), opts?.type);
    case 'json': {
      const out: Record<string, any[]> = {};
      for (const name of wb.SheetNames) out[name] = sheet_to_json(wb.Sheets[name] || {});
      return encodeText(JSON.stringify(out, null, 2), opts?.type);
    }
    case 'xlsx':
    case 'xlsm':
      return convertOutputType(writeXlsxSync(wb, opts), opts?.type);
    default:
      throw new Error(`Unsupported bookType ${bookType}`);
  }
}

export async function writeAsync(wb: WorkBook, opts?: WritingOptions): Promise<any> {
  const bt = opts?.bookType || 'xlsx';
  if (bt !== 'xlsx' && bt !== 'xlsm') return write(wb, opts);
  return convertOutputType(await writeXlsxAsync(wb, opts), opts?.type);
}

function bookTypeFor(filename: string, explicit?: BookType): BookType {
  if (explicit) return explicit;
  const ext = path.extname(filename).toLowerCase();
  if (ext === '.csv') return 'csv';
  if (ext === '.txt' || ext === '.tsv') return 'txt';
  if (ext === '.html' || ext === '.htm') return 'html';
  if (ext === '.json') return 'json';
  if (ext === '' || ext === '.xlsx') return 'xlsx';
  if (ext === '.xlsm') return 'xlsm';
  throw new Error(`Unsupported file extension "${ext}" (node-xls writes .xlsx, .csv, .txt, .html, .json)`);
}

export function writeFile(wb: WorkBook, filename: string, opts?: WritingOptions): void {
  fs.writeFileSync(filename, write(wb, { ...opts, bookType: bookTypeFor(filename, opts?.bookType), type: 'buffer' }));
}

export const writeFileXLSX = writeFile;

export async function writeFileAsync(wb: WorkBook, filename: string, opts?: WritingOptions): Promise<void> {
  const data = await writeAsync(wb, { ...opts, bookType: bookTypeFor(filename, opts?.bookType), type: 'buffer' });
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
  split_cell,

  // Sheet creation & conversion
  sheet_new,
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
  sheet_get_cell,
  sheet_set_array_formula,

  // Cells
  format_cell,
  cell_set_number_format,
  cell_set_hyperlink,
  cell_set_internal_link,
  cell_add_comment,

  // Workbook
  book_new,
  book_append_sheet,
  book_set_sheet_visibility,

  // Styling extensions
  set_cell_style,
  set_range_style,

  // Dates
  date_to_serial: dateToSerial,
  serial_to_date: serialToDate,

  // Pivot table extensions
  add_pivot_table,
  compute_pivot_table: computePivotData,

  consts: { SHEET_VISIBLE: 0, SHEET_HIDDEN: 1, SHEET_VERY_HIDDEN: 2 } as const,
};

const XLSX = {
  version,
  SSF,
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
