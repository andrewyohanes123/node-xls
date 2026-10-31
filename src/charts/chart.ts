/**
 * Native Excel charts (DrawingML c:chartSpace parts).
 *
 * Series reference worksheet ranges; their current values are embedded as
 * numCache/strCache so viewers render the chart without recalculating.
 * Child element order follows dml-chart.xsd exactly.
 */
import { CellObject, ChartAxisOptions, ChartSeries, Color, SheetChart, WorkSheet } from '../types.js';
import { WriteContext } from '../xlsx/context.js';
import { CONTENT_TYPES, NS, REL_TYPES, XML_HEADER, escapeAttr, escapeXml } from '../xlsx/xml.js';
import { absoluteRef, encode_cell, parseSheetRef } from '../utils/address.js';
import { format_cell } from '../utils/format.js';
import { dateToSerial } from '../utils/dates.js';
import { resolveAnchor, twoCellAnchorXml } from '../drawing/anchor.js';
import { normalizeRgb } from '../styles/index.js';

const THEME_SCHEME = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const AX_CAT = 500000001;
const AX_VAL = 500000002;

/** DrawingML color element for a Color (theme index -> scheme color, tint -> lumMod/lumOff). */
export function drawingColor(color: Color | undefined, fallbackAccent: number): string {
  if (typeof color === 'string') return `<a:srgbClr val="${normalizeRgb(color).slice(2)}"/>`;
  if (color?.rgb) return `<a:srgbClr val="${normalizeRgb(color.rgb).slice(2)}"/>`;
  const scheme = color?.theme !== undefined ? THEME_SCHEME[color.theme] ?? 'accent1' : `accent${(fallbackAccent % 6) + 1}`;
  const tint = color?.tint ?? 0;
  if (!tint) return `<a:schemeClr val="${scheme}"/>`;
  const mod = Math.round((tint > 0 ? 1 - tint : 1 + tint) * 100000);
  return `<a:schemeClr val="${scheme}"><a:lumMod val="${mod}"/>${tint > 0 ? `<a:lumOff val="${Math.round(tint * 100000)}"/>` : ''}</a:schemeClr>`;
}

interface ResolvedRef {
  f: string;
  cells: Array<CellObject | undefined>;
}

function resolveRef(ctx: WriteContext, ref: string, defaultSheet: string, what: string): ResolvedRef {
  let parsed;
  try {
    parsed = parseSheetRef(ref);
  } catch {
    throw new Error(`Chart ${what}: invalid range "${ref}"`);
  }
  const sheet = parsed.sheet ?? defaultSheet;
  const ws = ctx.wb.Sheets[sheet];
  if (!ws) throw new Error(`Chart ${what}: sheet "${sheet}" not found`);
  const cells: Array<CellObject | undefined> = [];
  const { s, e } = parsed.range;
  for (let r = s.r; r <= e.r; r++) for (let c = s.c; c <= e.c; c++) cells.push(ws[encode_cell({ r, c })]);
  return { f: absoluteRef(sheet, parsed.range), cells };
}

const isRef = (s: string) => /!\$?[A-Za-z]{1,3}\$?\d+/.test(s);

function numberOf(ctx: WriteContext, cell: CellObject | undefined): number | undefined {
  if (!cell) return undefined;
  if (typeof cell.v === 'number' && Number.isFinite(cell.v)) return cell.v;
  if (cell.v instanceof Date) return dateToSerial(cell.v, { date1904: ctx.date1904 });
  return undefined;
}

function numRefXml(ctx: WriteContext, r: ResolvedRef): string {
  const fmt = r.cells.find((c) => c?.z !== undefined)?.z;
  const pts = r.cells
    .map((c, i) => {
      const n = numberOf(ctx, c);
      return n === undefined ? '' : `<c:pt idx="${i}"><c:v>${n}</c:v></c:pt>`;
    })
    .join('');
  return `<c:numRef><c:f>${escapeXml(r.f)}</c:f><c:numCache><c:formatCode>${escapeXml(typeof fmt === 'string' ? fmt : 'General')}</c:formatCode><c:ptCount val="${r.cells.length}"/>${pts}</c:numCache></c:numRef>`;
}

function strRefXml(r: ResolvedRef): string {
  const pts = r.cells
    .map((c, i) => (c && c.t !== 'z' ? `<c:pt idx="${i}"><c:v>${escapeXml(format_cell(c))}</c:v></c:pt>` : ''))
    .join('');
  return `<c:strRef><c:f>${escapeXml(r.f)}</c:f><c:strCache><c:ptCount val="${r.cells.length}"/>${pts}</c:strCache></c:strRef>`;
}

/** Category/x data: numeric caches when every cell is numeric, string caches otherwise. */
function axisDataXml(ctx: WriteContext, tag: string, ref: string, sheet: string): string {
  const r = resolveRef(ctx, ref, sheet, tag);
  const numeric = r.cells.every((c) => c === undefined || numberOf(ctx, c) !== undefined);
  return `<c:${tag}>${numeric ? numRefXml(ctx, r) : strRefXml(r)}</c:${tag}>`;
}

function seriesTextXml(ctx: WriteContext, name: string | undefined, sheet: string): string {
  if (name === undefined) return '';
  if (isRef(name)) return `<c:tx>${strRefXml(resolveRef(ctx, name, sheet, 'series name'))}</c:tx>`;
  return `<c:tx><c:v>${escapeXml(name)}</c:v></c:tx>`;
}

function richText(text: string, bold = true): string {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr b="${bold ? 1 : 0}"/></a:pPr><a:r><a:rPr lang="en-US" b="${bold ? 1 : 0}"/><a:t>${escapeXml(
    text
  )}</a:t></a:r></a:p></c:rich></c:tx>`;
}

function titleXml(text: string | undefined): string {
  return text ? `<c:title>${richText(text)}<c:overlay val="0"/></c:title>` : '';
}

function dLblsXml(show: boolean | undefined, percent = false): string {
  if (!show) return '';
  return `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${percent ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${
    percent ? 1 : 0
  }"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

function seriesSpPr(type: SheetChart['type'], s: ChartSeries, i: number): string {
  if (type === 'pie' || type === 'doughnut') return '';
  const fill = drawingColor(s.color, i);
  if (type === 'line' || type === 'scatter') {
    return `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>${fill}</a:solidFill><a:round/></a:ln></c:spPr>`;
  }
  return `<c:spPr><a:solidFill>${fill}</a:solidFill></c:spPr>`;
}

function markerXml(s: ChartSeries, i: number, defaultOn: boolean): string {
  const on = s.marker ?? defaultOn;
  if (!on) return '<c:marker><c:symbol val="none"/></c:marker>';
  const fill = drawingColor(s.color, i);
  return `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill>${fill}</a:solidFill><a:ln><a:solidFill>${fill}</a:solidFill></a:ln></c:spPr></c:marker>`;
}

function seriesXml(ctx: WriteContext, chart: SheetChart, s: ChartSeries, i: number, sheet: string): string {
  const head = `<c:idx val="${i}"/><c:order val="${i}"/>${seriesTextXml(ctx, s.name, sheet)}${seriesSpPr(chart.type, s, i)}`;
  const values = resolveRef(ctx, s.values, sheet, 'values');
  const cat = s.categories ? axisDataXml(ctx, chart.type === 'scatter' ? 'xVal' : 'cat', s.categories, sheet) : '';
  const val = `<c:${chart.type === 'scatter' ? 'yVal' : 'val'}>${numRefXml(ctx, values)}</c:${chart.type === 'scatter' ? 'yVal' : 'val'}>`;
  switch (chart.type) {
    case 'column':
    case 'bar':
      return `<c:ser>${head}<c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
    case 'line':
      return `<c:ser>${head}${markerXml(s, i, false)}${cat}${val}<c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
    case 'scatter':
      return `<c:ser>${head}${markerXml(s, i, true)}${cat}${val}<c:smooth val="${s.smooth ? 1 : 0}"/></c:ser>`;
    default:
      return `<c:ser>${head}${cat}${val}</c:ser>`;
  }
}

function axisXml(kind: 'cat' | 'val', id: number, cross: number, pos: string, opts: ChartAxisOptions | undefined, extra: { gridDefault: boolean; crossBetween?: string }): string {
  const scaling = `<c:scaling><c:orientation val="minMax"/>${opts?.max !== undefined ? `<c:max val="${opts.max}"/>` : ''}${
    opts?.min !== undefined ? `<c:min val="${opts.min}"/>` : ''
  }</c:scaling>`;
  const grid = opts?.gridlines ?? extra.gridDefault ? '<c:majorGridlines/>' : '';
  const title = opts?.title ? `<c:title>${richText(opts.title, false)}<c:overlay val="0"/></c:title>` : '';
  const numFmt = opts?.numFmt ? `<c:numFmt formatCode="${escapeAttr(opts.numFmt)}" sourceLinked="0"/>` : kind === 'val' ? '<c:numFmt formatCode="General" sourceLinked="1"/>' : '';
  const shared =
    `<c:axId val="${id}"/>${scaling}<c:delete val="0"/><c:axPos val="${pos}"/>${grid}${title}${numFmt}` +
    `<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/>`;
  if (kind === 'cat') return `<c:catAx>${shared}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  return `<c:valAx>${shared}<c:crossBetween val="${extra.crossBetween ?? 'between'}"/></c:valAx>`;
}

function plotAreaXml(ctx: WriteContext, input: SheetChart, sheet: string): string {
  if (!input.series || input.series.length === 0) throw new Error('Chart needs at least one series');
  // 100% stacked charts show their value axis as percentages, like Excel
  const chart =
    input.grouping === 'percentStacked' && !input.yAxis?.numFmt ? { ...input, yAxis: { ...(input.yAxis ?? {}), numFmt: '0%' } } : input;
  const series = chart.series.map((s, i) => seriesXml(ctx, chart, s, i, sheet)).join('');
  const axIds = `<c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/>`;
  const dl = dLblsXml(chart.dataLabels);
  switch (chart.type) {
    case 'column':
    case 'bar': {
      const grouping = chart.grouping && chart.grouping !== 'standard' ? chart.grouping : 'clustered';
      const overlap = grouping === 'clustered' ? '' : '<c:overlap val="100"/>';
      const horizontal = chart.type === 'bar';
      return (
        `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}${dl}<c:gapWidth val="${
          chart.gapWidth ?? 150
        }"/>${overlap}${axIds}</c:barChart>` +
        axisXml('cat', AX_CAT, AX_VAL, horizontal ? 'l' : 'b', chart.xAxis, { gridDefault: false }) +
        axisXml('val', AX_VAL, AX_CAT, horizontal ? 'b' : 'l', chart.yAxis, { gridDefault: true })
      );
    }
    case 'line':
    case 'area': {
      const grouping = chart.grouping && chart.grouping !== 'clustered' ? chart.grouping : 'standard';
      const tag = chart.type === 'line' ? 'lineChart' : 'areaChart';
      const tail = chart.type === 'line' ? `<c:marker val="1"/>${axIds}` : axIds;
      return (
        `<c:${tag}><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}${dl}${tail}</c:${tag}>` +
        axisXml('cat', AX_CAT, AX_VAL, 'b', chart.xAxis, { gridDefault: false }) +
        axisXml('val', AX_VAL, AX_CAT, 'l', chart.yAxis, { gridDefault: true, crossBetween: chart.type === 'area' ? 'midCat' : 'between' })
      );
    }
    case 'scatter':
      return (
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${dl}${axIds}</c:scatterChart>` +
        axisXml('val', AX_CAT, AX_VAL, 'b', chart.xAxis, { gridDefault: false, crossBetween: 'midCat' }) +
        axisXml('val', AX_VAL, AX_CAT, 'l', chart.yAxis, { gridDefault: true, crossBetween: 'midCat' })
      );
    case 'pie':
      return `<c:pieChart><c:varyColors val="1"/>${series}${dLblsXml(chart.dataLabels, true)}<c:firstSliceAng val="0"/></c:pieChart>`;
    case 'doughnut':
      return `<c:doughnutChart><c:varyColors val="1"/>${series}${dLblsXml(chart.dataLabels, true)}<c:firstSliceAng val="0"/><c:holeSize val="${
        chart.holeSize ?? 50
      }"/></c:doughnutChart>`;
    default:
      throw new Error(`Unsupported chart type "${(chart as SheetChart).type}"`);
  }
}

export function buildChartXml(ctx: WriteContext, chart: SheetChart, sheet: string): string {
  const legend = chart.legend === false ? '' : `<c:legend><c:legendPos val="${chart.legend ?? 'r'}"/><c:overlay val="0"/></c:legend>`;
  return (
    `${XML_HEADER}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="${ctx.date1904 ? 1 : 0}"/><c:roundedCorners val="0"/>` +
    `<c:chart>${titleXml(chart.title)}<c:autoTitleDeleted val="${chart.title ? 0 : 1}"/><c:plotArea><c:layout/>${plotAreaXml(
      ctx,
      chart,
      sheet
    )}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    '</c:chartSpace>'
  );
}

/** Writes a chart part and returns its two-cell anchor for the drawing. */
export function writeChart(ctx: WriteContext, ws: WorkSheet, sheet: string, chart: SheetChart, drawingPart: string, id: number): string {
  const part = `xl/charts/chart${++ctx.seq.charts}.xml`;
  ctx.pkg.addPart(part, buildChartXml(ctx, chart, sheet), CONTENT_TYPES.chart);
  const rid = ctx.pkg.addRel(drawingPart, REL_TYPES.chart, part);
  const anchor = resolveAnchor(ws, chart.anchor, { width: 480, height: 288 });
  const frame =
    '<xdr:graphicFrame macro="">' +
    `<xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(chart.name ?? `Chart ${ctx.seq.charts}`)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart r:id="${rid}"/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame>';
  return twoCellAnchorXml(anchor, frame, 'twoCell');
}
