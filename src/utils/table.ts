import { WorkSheet, WorkBook } from '../types.js';
import { aoa_to_sheet } from './sheet.js';
import { book_new, book_append_sheet } from './book.js';

export function table_to_sheet(table: string | any, opts?: any): WorkSheet {
  let html = '';
  if (typeof table === 'string') {
    html = table;
  } else if (table && typeof table.outerHTML === 'string') {
    html = table.outerHTML;
  } else if (table && typeof table.innerHTML === 'string') {
    html = `<table>${table.innerHTML}</table>`;
  } else {
    return {};
  }

  const rows: string[][] = [];
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let trMatch: RegExpExecArray | null;

  while ((trMatch = trRegex.exec(html)) !== null) {
    const rowHtml = trMatch[1];
    const cellRegex = /<(?:td|th)[^>]*>([\s\S]*?)<\/(?:td|th)>/gi;
    const rowData: string[] = [];
    let cellMatch: RegExpExecArray | null;

    while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
      const cellContent = cellMatch[1]
        .replace(/<[^>]+>/g, '') // strip inner HTML tags
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&nbsp;/g, ' ')
        .trim();
      rowData.push(cellContent);
    }

    if (rowData.length > 0) {
      rows.push(rowData);
    }
  }

  // Convert types if numeric
  const converted = rows.map((row) =>
    row.map((val) => {
      const trimmed = val.trim();
      if (trimmed === '') return '';
      const num = Number(trimmed);
      if (!isNaN(num) && String(num) === trimmed) return num;
      return val;
    })
  );

  return aoa_to_sheet(converted, opts);
}

export function table_to_book(table: string | any, opts?: any): WorkBook {
  const ws = table_to_sheet(table, opts);
  const wb = book_new();
  book_append_sheet(wb, ws, opts?.sheet || 'Sheet1');
  return wb;
}
