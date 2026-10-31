import { SheetChart, WorkSheet } from '../types.js';
import { WriteContext } from '../xlsx/context.js';

/** Writes a chart part and returns its drawing anchor. */
export function writeChart(_ctx: WriteContext, _ws: WorkSheet, _sheet: string, _chart: SheetChart, _drawingPart: string, _id: number): string {
  throw new Error('Charts are not implemented yet');
}
