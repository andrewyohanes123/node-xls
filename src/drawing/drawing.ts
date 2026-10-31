import { WorkSheet } from '../types.js';
import { WriteContext } from '../xlsx/context.js';

/** Writes the sheet drawing (images + charts); returns the <drawing> element. */
export function writeDrawing(_ctx: WriteContext, ws: WorkSheet, _sheetPart: string): string {
  if (!ws['!images']?.length && !ws['!charts']?.length) return '';
  throw new Error('Images and charts are not implemented yet');
}
