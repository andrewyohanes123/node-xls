export type CellType = 'b' | 'n' | 'e' | 's' | 'd' | 'z';

export interface CellAddress {
  c: number;
  r: number;
}

export interface Range {
  s: CellAddress;
  e: CellAddress;
}

export type RangeSpec = string | Range;

// ---------------------------------------------------------------------------
// Styles
// ---------------------------------------------------------------------------

export interface ColorObject {
  /** Hex RGB ('FF0000', '#FF0000', 'F00') or ARGB ('80FF0000'). */
  rgb?: string;
  /** Theme color index (0 = lt1/background ... 4 = accent1 ... 9 = accent6). */
  theme?: number;
  /** Lighten (0..1) or darken (-1..0) a theme/indexed color. */
  tint?: number;
  /** Legacy indexed palette color. */
  indexed?: number;
  auto?: boolean;
}

export type Color = string | ColorObject;

export interface Font {
  name?: string;
  sz?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
  strike?: boolean;
  color?: Color;
  vertAlign?: 'superscript' | 'subscript';
  outline?: boolean;
  shadow?: boolean;
  family?: number;
  charset?: number;
  scheme?: 'major' | 'minor' | 'none';
}

export type FillPattern =
  | 'none'
  | 'solid'
  | 'gray125'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray0625'
  | 'darkHorizontal'
  | 'darkVertical'
  | 'darkDown'
  | 'darkUp'
  | 'darkGrid'
  | 'darkTrellis'
  | 'lightHorizontal'
  | 'lightVertical'
  | 'lightDown'
  | 'lightUp'
  | 'lightGrid'
  | 'lightTrellis';

export interface Fill {
  patternType?: FillPattern;
  /** Pattern (foreground) color. For solid fills this is the cell color. */
  fgColor?: Color;
  bgColor?: Color;
}

export type BorderStyle =
  | 'none'
  | 'thin'
  | 'medium'
  | 'dashed'
  | 'dotted'
  | 'thick'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface BorderSide {
  style?: BorderStyle;
  color?: Color;
}

export interface Border {
  top?: BorderSide;
  bottom?: BorderSide;
  left?: BorderSide;
  right?: BorderSide;
  diagonal?: BorderSide;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface Alignment {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  /** 0-90 counter-clockwise, 91-180 clockwise (Excel encoding), or 255 for vertical text. Negative values -1..-90 are converted. */
  textRotation?: number;
  indent?: number;
  shrinkToFit?: boolean;
  readingOrder?: number;
}

export interface CellProtection {
  /** Default true: cell is locked when the sheet is protected. */
  locked?: boolean;
  /** Hide the formula when the sheet is protected. */
  hidden?: boolean;
}

export interface CellStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  /** Number format code ('0.00%') or built-in format id (e.g. 14). */
  numFmt?: string | number;
  protection?: CellProtection;
}

// ---------------------------------------------------------------------------
// Cells
// ---------------------------------------------------------------------------

export interface Comment {
  /** Author. */
  a?: string;
  /** Text. */
  t: string;
}
/** @deprecated use Comment */
export type CommentAuthor = Comment;

/** SheetJS-compatible comment list; set `hidden: false` to always show the note. */
export type Comments = Comment[] & { hidden?: boolean };

export interface Hyperlink {
  /** URL, mailto:, file path, or internal reference starting with '#' ('#Sheet2!A1'). */
  Target: string;
  Tooltip?: string;
}

export interface RichTextRun {
  text: string;
  font?: Font;
}

export interface CellObject {
  /** Raw value. */
  v?: any;
  /** Formatted text. */
  w?: string;
  t?: CellType;
  /** Formula without leading '='. */
  f?: string;
  /** Range of the enclosing array formula (on every cell of the array). */
  F?: string;
  /** Dynamic-array formula flag (SheetJS); written as a legacy array formula. */
  D?: boolean;
  /** Style. */
  s?: CellStyle;
  /** Comments / notes. */
  c?: Comments;
  /** Hyperlink. */
  l?: Hyperlink;
  /** Number format (SheetJS convention). `s.numFmt` takes precedence. */
  z?: string | number;
  /** Rich text runs (written as a rich shared string; `v` holds the plain text). */
  richText?: RichTextRun[];
}

// ---------------------------------------------------------------------------
// Sheet-level structures
// ---------------------------------------------------------------------------

export interface ColProperties {
  /** Width in characters as stored by Excel. */
  width?: number;
  /** Width in characters (SheetJS 'wch'). */
  wch?: number;
  /** Width in pixels. */
  wpx?: number;
  hidden?: boolean;
  /** Outline (grouping) level 1-7 (SheetJS key). */
  level?: number;
  /** @deprecated alias of level */
  outlineLevel?: number;
  /** Default style for empty cells of the column. */
  s?: CellStyle;
}

export interface RowProperties {
  /** Height in points. */
  hpt?: number;
  /** Height in pixels. */
  hpx?: number;
  hidden?: boolean;
  /** Outline (grouping) level 1-7 (SheetJS key). */
  level?: number;
  /** @deprecated alias of level */
  outlineLevel?: number;
  /** Default style for empty cells of the row. */
  s?: CellStyle;
}

export interface AutoFilter {
  ref: string;
}

export interface Margins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  /** Excel paper size code (1 = Letter, 9 = A4). */
  paperSize?: number;
  /** Print scale 10-400 (%). */
  scale?: number;
  /** Fit to N pages wide (0 = automatic). Enables fit-to-page. */
  fitToWidth?: number;
  /** Fit to N pages tall (0 = automatic). Enables fit-to-page. */
  fitToHeight?: number;
  /** Scale to fit (default: on when fitToWidth/fitToHeight are given). */
  fitToPage?: boolean;
}

/**
 * Sheet protection (SheetJS-compatible). Flags map 1:1 to the OOXML
 * attributes: `true` means the action is LOCKED while protected.
 * Defaults (Excel): selectLockedCells/selectUnlockedCells/objects/scenarios
 * allowed; everything else locked.
 */
export interface SheetProtection {
  /** Plain-text password (hashed with SHA-512 on write). */
  password?: string;
  /** Existing password hash, preserved from files that were read. */
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
  /** Legacy 16-bit password hash (hex), preserved from files that were read. */
  passwordHash?: string;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  objects?: boolean;
  scenarios?: boolean;
}

export interface FreezePane {
  /** Number of frozen columns. */
  xSplit?: number;
  /** Number of frozen rows. */
  ySplit?: number;
  /** First visible cell of the scrolling pane (defaults to the split point). */
  topLeftCell?: string;
}

export interface SheetView {
  showGridLines?: boolean;
  showRowColHeaders?: boolean;
  showZeros?: boolean;
  rightToLeft?: boolean;
  /** Zoom percentage 10-400. */
  zoomScale?: number;
}

// --- Conditional formatting -------------------------------------------------

export interface CFValue {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';
  value?: number | string;
}

export type CFOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

export type CFTimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

export type CFIconSet =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters';

interface CFBase {
  stopIfTrue?: boolean;
  /** Evaluation order across the sheet (1 = first). Rules without one follow, in order. */
  priority?: number;
}

/** Differential style applied by a rule (font, fill, border, numFmt). */
export type CFStyle = Pick<CellStyle, 'font' | 'fill' | 'border' | 'numFmt'>;

export type ConditionalFormatRule =
  | (CFBase & { type: 'cellIs'; operator: CFOperator; formulae: Array<string | number>; style: CFStyle })
  | (CFBase & { type: 'expression'; formula: string; style: CFStyle })
  | (CFBase & { type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'; text: string; style: CFStyle })
  | (CFBase & { type: 'top10'; rank: number; bottom?: boolean; percent?: boolean; style: CFStyle })
  | (CFBase & { type: 'aboveAverage'; below?: boolean; equalAverage?: boolean; stdDev?: number; style: CFStyle })
  | (CFBase & {
      type:
        | 'duplicateValues'
        | 'uniqueValues'
        | 'containsBlanks'
        | 'notContainsBlanks'
        | 'containsErrors'
        | 'notContainsErrors';
      style: CFStyle;
    })
  | (CFBase & { type: 'timePeriod'; timePeriod: CFTimePeriod; style: CFStyle })
  | { type: 'colorScale'; cfvo: CFValue[]; colors: Color[]; priority?: number }
  | { type: 'dataBar'; color: Color; min?: CFValue; max?: CFValue; showValue?: boolean; priority?: number }
  | { type: 'iconSet'; iconSet: CFIconSet; cfvo?: CFValue[]; reverse?: boolean; showValue?: boolean; priority?: number };

export interface ConditionalFormat {
  /** Target range(s), e.g. 'B2:B20' or 'A1:A5 C1:C5'. */
  ref: string;
  rules: ConditionalFormatRule[];
}

// --- Data validation --------------------------------------------------------

export interface DataValidation {
  /** Target range(s), e.g. 'C2:C100'. */
  ref: string;
  /** 'any' accepts every value (useful for input prompts only). */
  type: 'any' | 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom';
  operator?: CFOperator;
  /** Inline list values (type 'list'). Joined length must stay under 255 characters. */
  values?: Array<string | number>;
  /** First formula / bound, or a range for lists ('$A$1:$A$9', 'Lists!$A$1:$A$9'). */
  formula1?: string | number | Date;
  formula2?: string | number | Date;
  /** Default true. */
  allowBlank?: boolean;
  /** Hide the in-cell dropdown arrow for lists. */
  hideDropDown?: boolean;
  /** Show the error alert on invalid input (default true). */
  showErrorMessage?: boolean;
  /** Show the input prompt when the cell is selected (default true). */
  showInputMessage?: boolean;
  prompt?: { title?: string; message: string };
  error?: { title?: string; message: string; style?: 'stop' | 'warning' | 'information' };
}

// --- Excel tables -----------------------------------------------------------

export type TableTotalsFunction = 'none' | 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom';

export interface TableColumn {
  /** Defaults to the header cell text. */
  name?: string;
  totalsRowFunction?: TableTotalsFunction;
  /** Formula for totalsRowFunction 'custom' (no leading '='). */
  totalsRowFormula?: string;
  totalsRowLabel?: string;
  /** Calculated-column formula Excel applies to new rows (cells still need their own formulas). */
  calculatedColumnFormula?: string;
}

export interface TableDefinition {
  /** Unique table name (letters, digits, underscores; no spaces). */
  name: string;
  /** Full table range including the header row (and totals row if enabled). */
  ref: string;
  displayName?: string;
  /** Default true. */
  headerRow?: boolean;
  /** Last row of `ref` is a totals row. */
  totalsRow?: boolean;
  columns?: TableColumn[];
  /** Built-in table style, e.g. 'TableStyleMedium2' (default), 'TableStyleLight9'; null for no style. */
  style?: string | null;
  /** Default true. */
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  /** Show filter buttons in the header (default true). */
  autoFilter?: boolean;
}

// --- Images & charts ----------------------------------------------------------

export type ImageType = 'png' | 'jpeg' | 'gif' | 'bmp' | 'tiff' | 'emf' | 'wmf';

export interface DrawingAnchor {
  /** Top-left cell (e.g. 'B2'). */
  from: string | CellAddress;
  /** Bottom-right cell (exclusive edge). When omitted, width/height are used. */
  to?: string | CellAddress;
  /** Size in pixels (used when `to` is omitted). */
  width?: number;
  height?: number;
  /** Pixel offsets inside the `from` cell. */
  offsetX?: number;
  offsetY?: number;
  /** Pixel offsets inside the `to` cell (with `to`). */
  toOffsetX?: number;
  toOffsetY?: number;
}

export interface SheetImage {
  /** Image bytes (Buffer/Uint8Array) or base64 string (data URIs accepted). */
  data: Uint8Array | ArrayBuffer | string;
  /** Detected from the bytes when omitted. */
  type?: ImageType;
  anchor: DrawingAnchor;
  name?: string;
  /** Alternative text. */
  description?: string;
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';

export interface ChartSeries {
  /** Series name: literal text or a cell reference ("'Data'!$B$1"). */
  name?: string;
  /** Category (x) range, e.g. "'Data'!$A$2:$A$6" or "Data!A2:A6", or literal values. */
  categories?: string | Array<string | number>;
  /** Value (y) range, or literal numbers. */
  values: string | number[];
  color?: Color;
  /** Line/scatter: show markers (default true for scatter, false for line). */
  marker?: boolean;
  /** Line: smooth curve. */
  smooth?: boolean;
}

export interface ChartAxisOptions {
  title?: string;
  /** Number format for axis labels. */
  numFmt?: string;
  min?: number;
  max?: number;
  /** Show major gridlines (default: value axis on, category axis off). */
  gridlines?: boolean;
}

export interface SheetChart {
  type: ChartType;
  /** column/bar/line/area grouping (default 'clustered' / 'standard'). */
  grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
  title?: string;
  series: ChartSeries[];
  /** Legend position, or false to hide (default 'r'). */
  legend?: 'r' | 'l' | 't' | 'b' | 'tr' | false;
  /** Show data labels with values. */
  dataLabels?: boolean;
  xAxis?: ChartAxisOptions;
  yAxis?: ChartAxisOptions;
  /** Doughnut hole size % (default 50). */
  holeSize?: number;
  /** Bar/column gap width % (default 150). */
  gapWidth?: number;
  anchor: DrawingAnchor;
  name?: string;
}

// ---------------------------------------------------------------------------
// Pivot tables
// ---------------------------------------------------------------------------

export type PivotSubtotal =
  | 'sum'
  | 'count'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'countNums'
  | 'stdDev'
  | 'stdDevp'
  | 'var'
  | 'varp';

export interface PivotDataField {
  name?: string;
  field: string;
  subtotal?: PivotSubtotal;
  /** Number format applied to the value cells. */
  numFmt?: string;
}

export interface PivotFilter {
  field: string;
  /** Single selected item (all items when omitted; null selects the "(blank)" item). Dates are serial numbers. */
  selected?: string | number | boolean | null;
}

export interface PivotTableDefinition {
  id?: number;
  name: string;
  sourceSheet: string;
  sourceRange: Range;
  targetSheet: string;
  targetCell: CellAddress;
  targetRange?: Range;
  rows: string[];
  cols: string[];
  values: PivotDataField[];
  filters: PivotFilter[];
  showRowTotals?: boolean;
  showColTotals?: boolean;
  /** Show subtotal rows for outer row fields when there are 2+ row fields (default true). */
  showSubtotals?: boolean;
  styleName?: string;
  renderCells?: boolean;
  /** Optional caption written above the pivot (outside the pivot area). */
  title?: string;
  /** Source given as a table or defined name (as read from a file); resolved to sourceSheet/sourceRange. */
  sourceName?: string;
}

export interface PivotTableOptions {
  name?: string;
  sourceSheet: string;
  sourceRange?: RangeSpec;
  targetSheet?: string;
  targetCell?: string | CellAddress;
  rows?: string[];
  cols?: string[];
  values?: (string | PivotDataField)[];
  filters?: (string | PivotFilter)[];
  showRowTotals?: boolean;
  showColTotals?: boolean;
  showSubtotals?: boolean;
  styleName?: string;
  renderCells?: boolean;
  title?: string;
}

// ---------------------------------------------------------------------------
// Worksheet & workbook
// ---------------------------------------------------------------------------

export interface WorkSheet {
  [cell: string]: CellObject | any;
  '!ref'?: string;
  '!cols'?: ColProperties[];
  '!rows'?: RowProperties[];
  '!merges'?: Range[];
  '!pivots'?: PivotTableDefinition[];
  '!autofilter'?: AutoFilter;
  '!protect'?: SheetProtection;
  '!margins'?: Margins;
  '!pageSetup'?: PageSetup;
  /** Freeze panes: top-left scrolling cell ('B2') or explicit splits. */
  '!freeze'?: string | FreezePane;
  '!views'?: SheetView;
  '!tabColor'?: Color;
  '!condfmt'?: ConditionalFormat[];
  '!dataValidation'?: DataValidation[];
  '!tables'?: TableDefinition[];
  '!images'?: SheetImage[];
  '!charts'?: SheetChart[];
}

export interface DefinedName {
  Name: string;
  /** Formula / reference, e.g. "Sheet1!$A$1:$B$10" (no leading '='). */
  Ref: string;
  /** Sheet index for a sheet-scoped name. */
  Sheet?: number;
  Comment?: string;
  Hidden?: boolean;
}

export interface WorkbookSheetProps {
  /** 0 = visible, 1 = hidden, 2 = very hidden. */
  Hidden?: 0 | 1 | 2;
  name?: string;
  CodeName?: string;
}

export interface WorkbookProtection {
  password?: string;
  /** Legacy 16-bit password hash (hex), preserved from files that were read. */
  passwordHash?: string;
  /** Existing password hash, preserved from files that were read. */
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
  /** Prevent adding/removing/renaming sheets (default true when protection is set). */
  lockStructure?: boolean;
  lockWindows?: boolean;
}

export interface WorkBook {
  SheetNames: string[];
  Sheets: { [sheet: string]: WorkSheet };
  /** Document properties (SheetJS keys: Title, Subject, Author, Manager, Company, Category, Keywords, Comments, LastAuthor, CreatedDate, ModifiedDate). */
  Props?: Record<string, any>;
  Custprops?: Record<string, any>;
  Workbook?: {
    Sheets?: WorkbookSheetProps[];
    Names?: DefinedName[];
    Views?: Array<{ RTL?: boolean; activeTab?: number }>;
    WBProps?: { date1904?: boolean; CodeName?: string; [k: string]: any };
    Protection?: WorkbookProtection;
  };
  PivotTables?: PivotTableDefinition[];
}

export type BookType = 'xlsx' | 'xlsm' | 'csv' | 'txt' | 'html' | 'json';
export type OutputType = 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

export interface WritingOptions {
  type?: OutputType;
  bookType?: BookType;
  compression?: boolean;
  Props?: Record<string, any>;
  cellStyles?: boolean;
  /** Sheet to export for single-sheet formats (csv/txt/html). Name or index. */
  sheet?: string | number;
  /** CSV/TXT output options (as sheet_to_csv). */
  FS?: string;
  RS?: string;
  strip?: boolean;
  blankrows?: boolean;
  forceQuotes?: boolean;
  skipHidden?: boolean;
  dateNF?: string;
  rawNumbers?: boolean;
}

export interface ParsingOptions {
  type?: OutputType;
  /** Keep formulae (default true). */
  cellFormula?: boolean;
  cellHTML?: boolean;
  /** Keep number format codes in cell.z (default true). */
  cellNF?: boolean;
  /** Keep cell styles in cell.s (default true). */
  cellStyles?: boolean;
  /** Generate formatted text in cell.w (default true). */
  cellText?: boolean;
  /** Return date-formatted cells as JS Dates with t:'d'. */
  cellDates?: boolean;
  /** Only read the first N rows of every sheet. */
  sheetRows?: number;
  /** Only parse these sheets (names or indexes). */
  sheets?: string | number | Array<string | number>;
  raw?: boolean;
  dense?: boolean;
  /** CSV field separator. */
  FS?: string;
}

export interface AOA2SheetOpts {
  dateNF?: string;
  cellDates?: boolean;
  sheetStubs?: boolean;
  nullError?: boolean;
  /** Treat Date values as UTC instead of local time. */
  UTC?: boolean;
  /** Convert Dates to 1904-system serials (for workbooks with WBProps.date1904). */
  date1904?: boolean;
  /** Origin cell or row; -1 appends below the existing range. */
  origin?: number | string | CellAddress;
}

export interface JSON2SheetOpts extends AOA2SheetOpts {
  header?: string[];
  skipHeader?: boolean;
}

export interface Sheet2JSONOpts {
  raw?: boolean;
  rawNumbers?: boolean;
  range?: number | string | Range;
  header?: 'A' | 1 | string[];
  dateNF?: string;
  defval?: any;
  blankrows?: boolean;
  skipHidden?: boolean;
  UTC?: boolean;
  /** Interpret date serials in the 1904 system (set automatically for sheets read from 1904 workbooks). */
  date1904?: boolean;
}

export interface Sheet2CSVOpts {
  FS?: string;
  RS?: string;
  dateNF?: string;
  strip?: boolean;
  blankrows?: boolean;
  skipHidden?: boolean;
  forceQuotes?: boolean;
  rawNumbers?: boolean;
}

export interface Sheet2HTMLOpts {
  id?: string;
  editable?: boolean;
  header?: string;
  footer?: string;
}
