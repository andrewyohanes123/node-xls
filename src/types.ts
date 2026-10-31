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

export interface ColProperties {
  wch?: number;
  wpx?: number;
  width?: number;
  hidden?: boolean;
  outlineLevel?: number;
}

export interface RowProperties {
  hpt?: number;
  hpx?: number;
  hidden?: boolean;
  outlineLevel?: number;
}

export interface ColorObject {
  rgb?: string;
  theme?: number;
  tint?: number;
  auto?: boolean;
}

export type Color = string | ColorObject;

export interface Font {
  name?: string;
  sz?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'single' | 'double';
  strike?: boolean;
  color?: Color;
  family?: number;
  scheme?: string;
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
  horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  textRotation?: number;
  indent?: number;
  shrinkToFit?: boolean;
  readingOrder?: number;
}

export interface CellStyle {
  font?: Font;
  fill?: Fill;
  border?: Border;
  alignment?: Alignment;
  numFmt?: string | number;
}

export interface CommentAuthor {
  a?: string;
  t: string;
}

export interface Hyperlink {
  Target: string;
  Tooltip?: string;
}

export interface CellObject {
  v?: any;
  w?: string;
  t?: CellType;
  f?: string;
  F?: string;
  s?: CellStyle;
  c?: CommentAuthor[];
  l?: Hyperlink;
  z?: string | number;
}

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
  numFmt?: string;
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
  filters: string[];
  showRowTotals?: boolean;
  showColTotals?: boolean;
  styleName?: string;
  renderCells?: boolean;
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
  filters?: string[];
  showRowTotals?: boolean;
  showColTotals?: boolean;
  styleName?: string;
  renderCells?: boolean;
}

export interface AutoFilter {
  ref: string;
}

export interface WorkSheet {
  [cell: string]: CellObject | any;
  '!ref'?: string;
  '!cols'?: ColProperties[];
  '!rows'?: RowProperties[];
  '!merges'?: Range[];
  '!pivots'?: PivotTableDefinition[];
  '!autofilter'?: AutoFilter;
  '!protect'?: any;
}

export interface WorkBook {
  SheetNames: string[];
  Sheets: { [sheet: string]: WorkSheet };
  Props?: Record<string, any>;
  Custprops?: Record<string, any>;
  Workbook?: {
    Views?: any[];
    Names?: any[];
    WBProps?: Record<string, any>;
  };
  PivotTables?: PivotTableDefinition[];
}

export type BookType = 'xlsx' | 'csv' | 'txt' | 'html' | 'json';
export type OutputType = 'base64' | 'binary' | 'buffer' | 'file' | 'array' | 'string';

export interface WritingOptions {
  type?: OutputType;
  bookType?: BookType;
  compression?: boolean;
  Props?: Record<string, any>;
  cellStyles?: boolean;
}

export interface ParsingOptions {
  type?: OutputType;
  cellFormula?: boolean;
  cellHTML?: boolean;
  cellNF?: boolean;
  cellStyles?: boolean;
  cellDates?: boolean;
  sheetRows?: number;
  raw?: boolean;
  dense?: boolean;
  FS?: string;
}

export interface AOA2SheetOpts {
  dateNF?: string;
  cellDates?: boolean;
  sheetStubs?: boolean;
  origin?: number | string | CellAddress;
}

export interface JSON2SheetOpts extends AOA2SheetOpts {
  header?: string[];
  skipHeader?: boolean;
}

export interface Sheet2JSONOpts {
  raw?: boolean;
  range?: number | string | Range;
  header?: 'A' | number | string[];
  dateNF?: string;
  defval?: any;
  blankrows?: boolean;
}

export interface Sheet2CSVOpts {
  FS?: string;
  RS?: string;
  dateNF?: string;
  strip?: boolean;
  blankrows?: boolean;
  skipHidden?: boolean;
  forceQuotes?: boolean;
}

export interface Sheet2HTMLOpts {
  id?: string;
  editable?: boolean;
  header?: string;
  footer?: string;
}
