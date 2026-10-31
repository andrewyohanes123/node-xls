import XLSX, { utils, WorkBook, WorkSheet, WritingOptions, ParsingOptions } from '../../src/index.js';
import { validateXlsx, XlsxReport } from './validate.js';

/** Workbook from [name, sheet] pairs. */
export function book(...sheets: Array<[string, WorkSheet]>): WorkBook {
  const wb = utils.book_new();
  for (const [name, ws] of sheets) utils.book_append_sheet(wb, ws, name);
  return wb;
}

export interface RoundTrip {
  buf: Uint8Array;
  /** Validator report: parts, SheetJS CE and ExcelJS views of the file. */
  rep: XlsxReport;
  /** The file read back with node-xls. */
  back: WorkBook;
}

/** Writes, validates (schema + independent readers) and reads back. */
export async function roundtrip(wb: WorkBook, opts?: WritingOptions, readOpts?: ParsingOptions): Promise<RoundTrip> {
  const buf: Uint8Array = XLSX.write(wb, { ...opts, type: 'array' });
  const rep = await validateXlsx(buf);
  const back = XLSX.read(buf, readOpts);
  return { buf, rep, back };
}

/** Round trip of a single sheet named 'S'. */
export async function roundtripSheet(ws: WorkSheet, readOpts?: ParsingOptions): Promise<RoundTrip & { ws: WorkSheet }> {
  const rt = await roundtrip(book(['S', ws]), undefined, readOpts);
  return { ...rt, ws: rt.back.Sheets.S };
}
