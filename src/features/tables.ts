/**
 * Excel tables (ListObjects): xl/tables/tableN.xml + <tableParts>.
 *
 * Excel requires header cells to hold the column names as strings and the
 * totals row to hold SUBTOTAL formulas; `tableCellOverrides` produces those
 * cells without mutating the user's worksheet.
 */
import { CellObject, Range, TableDefinition, TableTotalsFunction, WorkSheet } from '../types.js';
import { WriteContext } from '../xlsx/context.js';
import { CONTENT_TYPES, NS, REL_TYPES, XML_HEADER, attrs, escapeXml } from '../xlsx/xml.js';
import { decode_range, encode_cell, encode_range } from '../utils/address.js';
import { format_cell } from '../utils/format.js';

const SUBTOTAL_CODES: Record<Exclude<TableTotalsFunction, 'none' | 'custom'>, number> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};

const TABLE_NAME = /^[A-Za-z_\\][A-Za-z0-9_.\\]*$/;

interface ResolvedTable {
  def: TableDefinition;
  range: Range;
  names: string[];
}

function escapeStructuredName(name: string): string {
  return name.replace(/(['#[\]])/g, "'$1");
}

function resolve(ws: WorkSheet, def: TableDefinition): ResolvedTable {
  if (!TABLE_NAME.test(def.name) || /^[A-Za-z]{1,3}\d+$/.test(def.name) || /^[RrCc]\d*$/.test(def.name)) {
    throw new Error(`Invalid table name "${def.name}": use letters, digits and underscores, starting with a letter`);
  }
  const range = decode_range(def.ref.replace(/\$/g, ''));
  const headerRow = def.headerRow !== false;
  const width = range.e.c - range.s.c + 1;
  const minRows = (headerRow ? 1 : 0) + (def.totalsRow ? 1 : 0) + 1;
  if (range.e.r - range.s.r + 1 < minRows) throw new Error(`Table ${def.name} range ${def.ref} is too small`);

  const seen = new Set<string>();
  const names: string[] = [];
  for (let i = 0; i < width; i++) {
    let name = def.columns?.[i]?.name;
    if (name === undefined && headerRow) {
      const cell = ws[encode_cell({ c: range.s.c + i, r: range.s.r })] as CellObject | undefined;
      const text = cell ? format_cell(cell) : '';
      name = text.trim() === '' ? undefined : text;
    }
    const base = name ?? `Column${i + 1}`;
    let candidate = base;
    for (let n = 2; seen.has(candidate.toLowerCase()); n++) candidate = `${base}${n}`;
    seen.add(candidate.toLowerCase());
    names.push(candidate);
  }
  return { def, range, names };
}

/** Header/totals cells the table needs, keyed by cell reference. */
export function tableCellOverrides(ws: WorkSheet): Map<string, CellObject> {
  const out = new Map<string, CellObject>();
  for (const def of ws['!tables'] ?? []) {
    const { range, names } = resolve(ws, def);
    if (def.headerRow !== false) {
      names.forEach((name, i) => {
        const ref = encode_cell({ c: range.s.c + i, r: range.s.r });
        const cell = ws[ref] as CellObject | undefined;
        if (!cell || cell.t !== 's' || cell.v !== name || cell.f) out.set(ref, { t: 's', v: name, s: cell?.s, c: cell?.c, l: cell?.l });
      });
    }
    if (def.totalsRow) {
      (def.columns ?? []).forEach((col, i) => {
        if (!col) return;
        const ref = encode_cell({ c: range.s.c + i, r: range.e.r });
        const cell = ws[ref] as CellObject | undefined;
        if (col.totalsRowFunction === 'custom') {
          if (!col.totalsRowFormula) throw new Error(`Table ${def.name}: column ${names[i]} has totalsRowFunction "custom" but no totalsRowFormula`);
          out.set(ref, { t: 'n', f: col.totalsRowFormula.replace(/^=/, ''), s: cell?.s, c: cell?.c, l: cell?.l });
        } else if (col.totalsRowFunction && col.totalsRowFunction !== 'none') {
          const code = SUBTOTAL_CODES[col.totalsRowFunction];
          out.set(ref, { t: 'n', f: `SUBTOTAL(${code},${def.name}[${escapeStructuredName(names[i])}])`, s: cell?.s, c: cell?.c, l: cell?.l });
        } else if (col.totalsRowLabel !== undefined && !cell) {
          out.set(ref, { t: 's', v: col.totalsRowLabel });
        }
      });
    }
  }
  return out;
}

function overlaps(a: Range, b: Range): boolean {
  return a.s.r <= b.e.r && b.s.r <= a.e.r && a.s.c <= b.e.c && b.s.c <= a.e.c;
}

/** Writes table parts for a sheet; returns the <tableParts> element. */
export function writeTables(ctx: WriteContext, ws: WorkSheet, sheetName: string, sheetPart: string): string {
  const defs = ws['!tables'];
  if (!defs || defs.length === 0) return '';
  const resolved = defs.map((d) => resolve(ws, d));
  const rids: string[] = [];

  resolved.forEach((t, i) => {
    const key = t.def.name.toLowerCase();
    if (ctx.tableNames.has(key)) throw new Error(`Duplicate table name "${t.def.name}"`);
    ctx.tableNames.add(key);
    for (let j = 0; j < i; j++) {
      if (overlaps(t.range, resolved[j].range)) throw new Error(`Tables ${t.def.name} and ${resolved[j].def.name} overlap on ${sheetName}`);
    }
    for (const m of ws['!merges'] ?? []) {
      if (overlaps(t.range, m)) throw new Error(`Table ${t.def.name} overlaps merged cells ${encode_range(m)}`);
    }
    const af = ws['!autofilter']?.ref;
    if (af && overlaps(t.range, decode_range(af.replace(/\$/g, '')))) {
      throw new Error(`Table ${t.def.name} overlaps the sheet autofilter; remove !autofilter`);
    }

    const id = ++ctx.seq.tables;
    const { def, range, names } = t;
    const headerRow = def.headerRow !== false;
    const dataEnd = def.totalsRow ? range.e.r - 1 : range.e.r;
    const autoFilter =
      headerRow && def.autoFilter !== false
        ? `<autoFilter ref="${encode_range({ s: range.s, e: { c: range.e.c, r: dataEnd } })}"/>`
        : '';
    const columns = names
      .map((name, i) => {
        const col = def.columns?.[i];
        const fn = col?.totalsRowFunction && col.totalsRowFunction !== 'none' ? col.totalsRowFunction : undefined;
        const children =
          (col?.calculatedColumnFormula ? `<calculatedColumnFormula>${escapeXml(col.calculatedColumnFormula.replace(/^=/, ''))}</calculatedColumnFormula>` : '') +
          (def.totalsRow && fn === 'custom' && col?.totalsRowFormula ? `<totalsRowFormula>${escapeXml(col.totalsRowFormula.replace(/^=/, ''))}</totalsRowFormula>` : '');
        const a = attrs([
          ['id', i + 1],
          // column names are ST_Xstring: line breaks are stored as _x000a_ (as Excel does)
          ['name', name.replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1').replace(/\n/g, '_x000a_').replace(/\r/g, '_x000d_')],
          ['totalsRowFunction', def.totalsRow ? fn : undefined],
          ['totalsRowLabel', def.totalsRow && !fn ? col?.totalsRowLabel : undefined],
        ]);
        return children ? `<tableColumn${a}>${children}</tableColumn>` : `<tableColumn${a}/>`;
      })
      .join('');
    const styleInfo = `<tableStyleInfo${attrs([
      ['name', def.style ?? 'TableStyleMedium2'],
      ['showFirstColumn', def.showFirstColumn ? '1' : '0'],
      ['showLastColumn', def.showLastColumn ? '1' : '0'],
      ['showRowStripes', def.showRowStripes === false ? '0' : '1'],
      ['showColumnStripes', def.showColumnStripes ? '1' : '0'],
    ])}/>`;
    const xml = `${XML_HEADER}<table xmlns="${NS.main}"${attrs([
      ['id', id],
      ['name', def.name],
      ['displayName', def.displayName ?? def.name],
      ['ref', encode_range(range)],
      ['headerRowCount', headerRow ? undefined : '0'],
      ['totalsRowCount', def.totalsRow ? 1 : undefined],
      ['totalsRowShown', def.totalsRow ? undefined : '0'],
    ])}>${autoFilter}<tableColumns count="${names.length}">${columns}</tableColumns>${styleInfo}</table>`;
    const part = `xl/tables/table${id}.xml`;
    ctx.pkg.addPart(part, xml, CONTENT_TYPES.table);
    rids.push(ctx.pkg.addRel(sheetPart, REL_TYPES.table, part));
  });

  return `<tableParts count="${rids.length}">${rids.map((r) => `<tablePart r:id="${r}"/>`).join('')}</tableParts>`;
}
