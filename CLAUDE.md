# CLAUDE.md — node-xls contributor guide

node-xls reads and writes `.xlsx` with a SheetJS CE–compatible API and adds what SheetJS sells as Pro: styles, rich text, conditional formatting, data validation, tables, images, charts and pivot tables. TypeScript, Node ≥ 18, runtime deps: `fflate` (zip), `fast-xml-parser` (reader). Number formatting is SheetJS CE's SSF, vendored in `src/vendor/ssf.ts` (Apache-2.0, `THIRD_PARTY_NOTICES.md`); the npm `ssf` 0.11.2 rounds times wrongly, so don't switch back to it.

## Commands

| Command | Purpose |
|---|---|
| `npm test` | Full vitest suite (~360 tests). Needs the XSDs from `npm run schemas` and `xmllint`. |
| `npm run test:tz` | Suite under UTC, Asia/Jakarta, America/New_York, Pacific/Kiritimati (date handling). |
| `npm run schemas` | Downloads the ECMA-376 Transitional XSDs to `.cache/ecma/xsd` (gitignored). |
| `npm run build` | Clean `dist/` and compile with `tsc` (CommonJS + .d.ts). |
| `npm run showcase` | Build, then write one workbook per feature area to `out/showcase/`. |
| `npm run render [files]` | LibreOffice → PDF → PNG in `out/render/`; exits non-zero if LibreOffice can't open a file. |
| `npm run test:lo` | `showcase` + `render` (LibreOffice smoke test). |
| `npm run inspect -- <file>` | List the parts of an xlsx and pretty-print every XML part to `out/inspect/<name>/`. |
| `npm run excel-check` | Focused files + `CHECKLIST.md` in `out/excel-check/` for manual checks in Microsoft Excel. |

Run one test file: `npx vitest run test/pivot.test.ts`. Run one test: `npx vitest run -t "freeze panes"`.

## Architecture

```
src/index.ts              public API: read/write/readFile/writeFile(+Async), utils, SSF
src/types.ts              all public types (cell/sheet/workbook model, feature options)
src/xlsx/
  writer.ts               orchestrates a package: sheets, pivots, theme, styles, sst, workbook
  package.ts              PackageWriter: parts, [Content_Types].xml, auto rIds, relative targets; resolveTarget()
  worksheet.ts            sheet XML; sections emitted in WORKSHEET_CHILD_ORDER (schema order)
  workbook.ts             workbook.xml (views, sheets, defined names, calcPr, pivotCaches), docProps, sheet-name rules
  sst.ts                  shared strings (plain + rich runs)
  xml.ts                  escaping (escapeXml / escapeAttr / ST_Xstring), namespaces, rel + content types
  context.ts              WriteContext shared by part writers (styles, sst, seq counters, media)
  reader.ts               reader; parts located through relationships (rels.ts), never hard-coded
  read-features.ts        parsers for CF, data validation, tables, comments
src/styles/
  index.ts                StyleBuilder (fonts/fills/borders/numFmts/cellXfs/dxfs), color helpers
  parser.ts               styles.xml -> CellStyle[] / dxfs
  theme.ts                default Office theme part (theme colors resolve against it)
  css.ts                  CellStyle <-> inline CSS (sheet_to_html / table_to_sheet)
src/features/             comments (+VML), hyperlinks, conditional, validation, tables, protection
src/drawing/              drawing part, anchors (px -> two-cell anchors), images, drawing reader
src/charts/               chart part writer (dml-chart order) and reader
src/pivot/
  source.ts               source range -> unique field names + typed records
  layout.ts               layout engine: items, axes (items/subtotals/grand), values, geometry
  render.ts               layout -> worksheet cells (Excel tabular form)
  openxml.ts              layout -> pivotCacheDefinition / records / pivotTableDefinition
  compute.ts              utils.compute_pivot_table (summary API)
  parser.ts               pivot definitions back to PivotTableDefinition
src/utils/                address, sheet utils (SheetJS semantics), book utils, dates, format (SSF), formula, units, table (HTML import)
src/csv/                  CSV/TSV parse + CSV/HTML output
src/vendor/ssf.ts         SheetJS SSF (vendored, @ts-nocheck); only touch it to sync with upstream
test/helpers/validate.ts  the validation oracle (see Testing)
scripts/                  showcase, render, inspect, excel-check, fetch-schemas, png encoder
```

## Data model conventions

- SheetJS shapes everywhere: `cell = { t, v, w, f, F, z, s, l, c }`, sheet settings under `!` keys. Extensions also use `!` keys: `!freeze`, `!views`, `!tabColor`, `!condfmt`, `!dataValidation`, `!tables`, `!images`, `!charts`, `!pageSetup`, `!pivots`. Cells add `richText`.
- `cell.s.numFmt` takes precedence over `cell.z`. Date values without a format get built-in 14.
- **Dates use local calendar fields** (`utils/dates.ts`, matches SheetJS). ISO strings in `t:'d'` cells use their literal fields (`isoToSerial`). `new Date('2024-02-29')` is UTC midnight and lands on the wrong day in negative-offset timezones, so never use it for this.
- When behaviour overlaps SheetJS CE, match SheetJS exactly and prove it with a differential test (`test/sheet-utils.test.ts`). Deliberate deviations are listed in the README under "Differences from SheetJS CE".
- Never mutate the user's workbook while writing (e.g. table header overrides are computed in `tableCellOverrides`, not written back).

## OOXML rules that bit us (keep them)

1. **Element order is schema order.** CT_Worksheet, CT_Stylesheet, CT_Workbook, CT_Chart… are sequences. Use the order constants and builders; the XSD tests catch violations. Example: `autoFilter` comes before `mergeCells`.
2. **Text escaping:** use `escapeXml` for text nodes and `escapeAttr` for attributes (escapeXml does not escape `"`). Cell/run/comment text goes through `encodeXstring`: XML-illegal chars become `_xHHHH_`, a literal `_xHHHH_` becomes `_x005F_xHHHH_`, and CR becomes `_x000D_` (parsers normalise CR).
3. **Reader parsers must use `parseTagValue: false, trimValues: false`.** Otherwise `"007"` becomes `7` and padding is lost.
4. **Relationship targets** can be relative or absolute (`/xl/...`). Always go through `resolveTarget`/`readRels`.
5. **Styles:** fonts/borders/cellXfs index 0 and fills 0 (`none`) / 1 (`gray125`) are reserved defaults. Conditional formats use `dxfs`, where a solid fill colour is stored in `bgColor`.
6. **Theme colors** need the theme part (`styles/theme.ts`); without it LibreOffice renders theme fills black.
7. **Rich-text runs** with a font must carry complete run properties (size/name/…), otherwise apps fall back to a default serif font.
8. **Comments** need both `commentsN.xml` (text as `<r>` runs; ExcelJS ignores a bare `<t>`) and the legacy VML drawing.
9. **Data validation `allowBlank`** defaults to false in the schema (absent = blanks not allowed).
10. **Pivots:** shared items only for axis/page fields, records use `<x v>` for them; `pivotField` items are in display order, and `rowItems`/`colItems` index those items (not the shared items). More than one data field adds field `x="-2"` to `colFields`. `location` must equal header rows + item rows; page fields sit above `location`, so nothing may be rendered there. Data field names must not equal a source field name. Everything comes from `pivot/layout.ts`, so the renderer and the XML can't disagree.
11. **Excel limits enforced on write** (README "Excel limits" has the full list): 32,767 chars per cell/comment, 8,192 per formula, sheet names (31 chars, no `[]:*?/\` or control chars, not `History`, case-insensitive unique), name grammar for defined/table names (`checkExcelName`), every range inside A1:XFD1048576 (`checkRangeInGrid`/`checkSqref`), outline levels ≤ 7, style value ranges and ≤ 65,430 cellXfs, DV error ≤ 224 chars, inline list ≤ 255 chars, no overlapping merges/tables/pivots. New features must throw rather than write values Excel rejects.
12. **Rich-text runs** don't inherit on/off flags (bold, italic, …) from the cell font, only name/size/colour (`sst.ts`).
13. **Number format ids 5–8, 27–36, 50–82 without a code** are locale-dependent; we resolve them through SSF's implicit table exactly as SheetJS does (`builtinFormatCode`).

## Testing rules

- **Every test that writes a file must go through `validateXlsx`** (or the `roundtrip` / `roundtripSheet` helpers in `test/helpers/util.ts`). It checks:
  - content types, dangling relationships, XML well-formedness;
  - shared-string and style index ranges, row/cell order;
  - **ECMA-376 XSD validation** via xmllint (Markup Compatibility is stripped for third-party files);
  - that **SheetJS CE and ExcelJS** can load the file.
  It returns both parsed views, so you can assert on what an independent reader sees.
- Assert features through an **independent reader** (ExcelJS for styles/CF/validation/images/tables, SheetJS CE for values/links/comments/names), not only through node-xls' own reader. The original tests only round-tripped through our own reader and missed every corruption bug.
- `test/validator.test.ts` proves the oracle accepts SheetJS/ExcelJS output and rejects bad packages. Keep it green before trusting other failures.
- `test/review-fixes.test.ts` and `test/review2.test.ts` hold the regression tests from the two independent reviews (numbered by finding). Each one was checked to fail on the code before its fix; keep it that way when adding to them.
- Pivot tests use `checkPivotParts`, a consistency check across cache/records/table/location, as a stand-in for Excel's own load-time validation.
- Date-related changes must pass `npm run test:tz`.
- Visual changes: `npm run showcase && npm run render`, then look at the PNGs. Excel itself (not on Linux) is the final judge for pivots and charts; regenerate `npm run excel-check` and ask for a check.

## Adding a feature (checklist)

1. Types in `src/types.ts` (SheetJS-style `!` key or cell field).
2. Writer in `src/features/` (or `drawing/`, `charts/`), hooked into `worksheet.ts` at its schema position; register parts with `ctx.pkg.addPart(path, xml, CONTENT_TYPES.x)` and relationships with `ctx.pkg.addRel(source, REL_TYPES.x, target)`.
3. Reader support so read → write round-trips.
4. Tests: XSD-valid, independent reader sees it, round trip, error cases.
5. A showcase section plus a render check; README section.

## Known gaps / open items

- Pivot tables and charts are validated against the schema, an internal consistency checker, ExcelJS and LibreOffice (which recomputes pivots identically), but have **not yet been confirmed in Microsoft Excel**. `out/excel-check/CHECKLIST.md` lists what to check. `pivot-8-ab-no-items.xlsx` is an A/B diagnostic in case Excel rejects the `rowItems`/`colItems`.
- Not implemented: legacy formats (xls/xlsb/ods), encryption, formula evaluation, VBA, threaded comments, sparklines, combo/secondary-axis charts, compact/outline pivot layouts, calculated pivot fields, `XLSX.stream`, pictures inside group shapes (`xdr:grpSp`, dropped on read), preserving unmodelled parts on round trip.
