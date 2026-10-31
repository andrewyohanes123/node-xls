import { WorkBook, WritingOptions } from '../types.js';
import { PackageWriter } from './package.js';
import { StyleBuilder } from '../styles/index.js';
import { SharedStringsBuilder } from './sst.js';

/** State shared by all part writers while building one package. */
export interface WriteContext {
  wb: WorkBook;
  opts?: WritingOptions;
  pkg: PackageWriter;
  styles: StyleBuilder;
  sst: SharedStringsBuilder;
  date1904: boolean;
  /** Index of the active (selected) sheet. */
  activeTab: number;
  /** Sequence numbers for numbered parts. */
  seq: { comments: number; tables: number; drawings: number; charts: number; images: number; vmlBlocks: number };
  /** Workbook-wide table names (case-insensitive) and ids. */
  tableNames: Set<string>;
  hasFormulas: boolean;
  /** Hidden defined names generated while writing sheets (e.g. _xlnm._FilterDatabase). */
  autoNames: Array<{ name: string; sheet: number; ref: string }>;
  /** Media parts by content hash (identical images are stored once). */
  media: Map<string, string>;
}

export function createContext(wb: WorkBook, opts?: WritingOptions): WriteContext {
  return {
    wb,
    opts,
    pkg: new PackageWriter(),
    styles: new StyleBuilder(),
    sst: new SharedStringsBuilder(),
    date1904: Boolean(wb.Workbook?.WBProps?.date1904),
    activeTab: 0,
    seq: { comments: 0, tables: 0, drawings: 0, charts: 0, images: 0, vmlBlocks: 0 },
    tableNames: new Set(),
    hasFormulas: false,
    autoNames: [],
    media: new Map(),
  };
}
