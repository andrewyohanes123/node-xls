import { WorkBook, WorkSheet, ParsingOptions, WritingOptions } from '../types.js';
import { sheet_to_csv, aoa_to_sheet } from '../utils/sheet.js';
import { book_new, book_append_sheet } from '../utils/book.js';

export function parseCsv(text: string, fs = ','): string[][] {
  const lines: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let insideQuotes = false;

  let i = 0;
  const len = text.length;

  while (i < len) {
    const char = text[i];

    if (insideQuotes) {
      if (char === '"') {
        if (i + 1 < len && text[i + 1] === '"') {
          // Escaped quote
          currentField += '"';
          i += 2;
          continue;
        } else {
          insideQuotes = false;
          i++;
          continue;
        }
      } else {
        currentField += char;
        i++;
        continue;
      }
    } else {
      if (char === '"') {
        insideQuotes = true;
        i++;
        continue;
      }

      if (char === fs) {
        currentRow.push(currentField);
        currentField = '';
        i++;
        continue;
      }

      if (char === '\r') {
        if (i + 1 < len && text[i + 1] === '\n') {
          i++;
        }
        currentRow.push(currentField);
        lines.push(currentRow);
        currentRow = [];
        currentField = '';
        i++;
        continue;
      }

      if (char === '\n') {
        currentRow.push(currentField);
        lines.push(currentRow);
        currentRow = [];
        currentField = '';
        i++;
        continue;
      }

      currentField += char;
      i++;
    }
  }

  if (currentField.length > 0 || currentRow.length > 0) {
    currentRow.push(currentField);
    lines.push(currentRow);
  }

  return lines;
}

export function readCsv(data: string, opts?: ParsingOptions): WorkBook {
  const fs = opts?.FS || (data.includes('\t') && !data.includes(',') ? '\t' : ',');
  const rows = parseCsv(data, fs);

  // Convert types if numeric/boolean
  const converted = rows.map((row) =>
    row.map((val) => {
      const trimmed = val.trim();
      if (trimmed === '') return '';
      if (trimmed === 'true') return true;
      if (trimmed === 'false') return false;
      const num = Number(trimmed);
      if (!isNaN(num) && String(num) === trimmed) return num;
      return val;
    })
  );

  const ws = aoa_to_sheet(converted);
  const wb = book_new();
  book_append_sheet(wb, ws, 'Sheet1');
  return wb;
}

export function writeCsv(wb: WorkBook, opts?: WritingOptions): string {
  const firstSheetName = wb.SheetNames[0] || 'Sheet1';
  const ws = wb.Sheets[firstSheetName] || {};
  return sheet_to_csv(ws, {
    FS: opts?.bookType === 'txt' ? '\t' : ',',
    blankrows: true,
  });
}
