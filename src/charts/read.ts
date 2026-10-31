import { XML_READ_OPTIONS } from '../xlsx/xml.js';
/**
 * Parses a chart part into a SheetChart (type, grouping, title, series
 * references, legend, axes options). Formatting beyond series colors is not
 * modelled; full-fidelity chart editing is out of scope.
 */
import { strFromU8 } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { ChartAxisOptions, ChartSeries, Color, DrawingAnchor, SheetChart } from '../types.js';

const parser = new XMLParser({
  ...XML_READ_OPTIONS,
  isArray: (n) => ['ser', 'p', 'r', 'catAx', 'valAx', 'dateAx'].includes(n),
});

const THEME_SCHEME = ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const val = (n: any) => n?.['@_val'];
/** CT_Boolean: present without val means true; 'true'/'1' are true. */
const flag = (n: any): boolean | undefined => (n === undefined ? undefined : n?.['@_val'] === undefined || n['@_val'] === '1' || n['@_val'] === 'true');

function richText(tx: any): string | undefined {
  const paragraphs = tx?.rich?.p;
  if (!paragraphs) return undefined;
  return paragraphs
    .map((p: any) => (p.r ?? []).map((r: any) => (typeof r.t === 'object' ? (r.t['#text'] ?? '') : String(r.t ?? ''))).join(''))
    .join('\n');
}

function colorOf(spPr: any): Color | undefined {
  const fill = spPr?.solidFill ?? spPr?.ln?.solidFill;
  if (!fill) return undefined;
  if (fill.srgbClr) return { rgb: String(val(fill.srgbClr)).toUpperCase() };
  if (fill.schemeClr) {
    const idx = THEME_SCHEME.indexOf(val(fill.schemeClr));
    if (idx < 0) return undefined;
    // lumMod/lumOff as written for tints (see drawingColor): darken = lumMod only, lighten = lumMod + lumOff
    const mod = val(fill.schemeClr.lumMod);
    const off = val(fill.schemeClr.lumOff);
    if (mod === undefined) return { theme: idx };
    const tint = off !== undefined ? Number(off) / 100000 : Number(mod) / 100000 - 1;
    return { theme: idx, tint: Math.round(tint * 1000) / 1000 };
  }
  return undefined;
}

/** Series data: a range reference, or the literal points of numLit/strLit. */
function refOf(node: any): string | Array<string | number> | undefined {
  const f = node?.numRef?.f ?? node?.strRef?.f;
  if (f !== undefined) return typeof f === 'object' ? f['#text'] : String(f);
  const lit = node?.numLit ?? node?.strLit;
  if (!lit) return undefined;
  const count = Number(val(lit.ptCount) ?? 0);
  const out: Array<string | number> = Array.from({ length: count }, () => (node.numLit ? NaN : ''));
  for (const pt of [].concat(lit.pt ?? [])) {
    const raw = (pt as any).v;
    const text = typeof raw === 'object' ? (raw['#text'] ?? '') : String(raw ?? '');
    out[Number((pt as any)['@_idx'] ?? 0)] = node.numLit ? Number(text) : text;
  }
  return out;
}

function axisOf(ax: any, kind: 'cat' | 'val'): ChartAxisOptions | undefined {
  if (!ax) return undefined;
  const o: ChartAxisOptions = {};
  const title = richText(ax.title?.tx);
  if (title) o.title = title;
  const linked = ax.numFmt?.['@_sourceLinked'];
  if (ax.numFmt?.['@_formatCode'] && !(linked === '1' || linked === 'true')) o.numFmt = ax.numFmt['@_formatCode'];
  if (val(ax.scaling?.min) !== undefined) o.min = Number(val(ax.scaling.min));
  if (val(ax.scaling?.max) !== undefined) o.max = Number(val(ax.scaling.max));
  // node-xls defaults: gridlines on value axes, off on category axes
  const grid = ax.majorGridlines !== undefined;
  if (grid !== (kind === 'val')) o.gridlines = grid;
  return Object.keys(o).length ? o : undefined;
}

export function readChartPart(files: Record<string, Uint8Array>, part: string, anchor: DrawingAnchor): SheetChart | undefined {
  const bytes = files[part];
  if (!bytes) return undefined;
  const chart = parser.parse(strFromU8(bytes)).chartSpace?.chart;
  const plot = chart?.plotArea;
  if (!plot) return undefined;

  let type: SheetChart['type'] | undefined;
  let node: any;
  for (const [tag, t] of [
    ['barChart', 'column'],
    ['lineChart', 'line'],
    ['areaChart', 'area'],
    ['pieChart', 'pie'],
    ['doughnutChart', 'doughnut'],
    ['scatterChart', 'scatter'],
  ] as const) {
    if (plot[tag]) {
      type = t;
      node = Array.isArray(plot[tag]) ? plot[tag][0] : plot[tag];
      break;
    }
  }
  if (!type || !node) return undefined;
  if (type === 'column' && val(node.barDir) === 'bar') type = 'bar';

  const series: ChartSeries[] = (node.ser ?? []).map((s: any) => {
    const values = refOf(type === 'scatter' ? s.yVal : s.val);
    const out: ChartSeries = { values: Array.isArray(values) ? values.map(Number) : (values ?? '') };
    const name = s.tx?.strRef?.f ?? s.tx?.v;
    if (name !== undefined) {
      const text = typeof name === 'object' ? name['#text'] : String(name);
      // LibreOffice writes literal names as a quoted "formula": <c:strRef><c:f>"Plan"</c:f>
      const literal = /^"((?:[^"]|"")*)"$/.exec(text);
      out.name = literal ? literal[1].replace(/""/g, '"') : text;
    }
    const cat = refOf(type === 'scatter' ? s.xVal : s.cat);
    if (cat) out.categories = cat;
    const color = colorOf(s.spPr);
    if (color) out.color = color;
    if (type === 'line' || type === 'scatter') {
      const symbol = val(s.marker?.symbol);
      if (symbol !== undefined) out.marker = symbol !== 'none';
      if (flag(s.smooth)) out.smooth = true;
    }
    return out;
  });

  const out: SheetChart = { type, series, anchor };
  const grouping = val(node.grouping);
  if (grouping && !(grouping === 'clustered' && (type === 'column' || type === 'bar')) && !(grouping === 'standard' && (type === 'line' || type === 'area'))) {
    out.grouping = grouping;
  }
  const title = richText(chart.title?.tx);
  if (title) out.title = title;
  out.legend = chart.legend ? (val(chart.legend.legendPos) ?? 'r') : false;
  const shown = (d: any) => d && (flag(d.showVal) || flag(d.showPercent));
  if (shown(node.dLbls) || (node.ser ?? []).some((s: any) => shown(s.dLbls))) out.dataLabels = true;
  if (type === 'doughnut' && val(node.holeSize) !== undefined) out.holeSize = Number(val(node.holeSize));
  if ((type === 'column' || type === 'bar') && val(node.gapWidth) !== undefined && Number(val(node.gapWidth)) !== 150) out.gapWidth = Number(val(node.gapWidth));

  // x = category axis, y = value axis (by type, not position: bar charts swap positions);
  // scatter charts have two value axes, told apart by position
  const catAxes = [...(plot.catAx ?? []), ...(plot.dateAx ?? [])];
  const valAxes = plot.valAx ?? [];
  let xAxis: ChartAxisOptions | undefined;
  let yAxis: ChartAxisOptions | undefined;
  if (type === 'scatter') {
    const horizontalAx = valAxes.find((a: any) => ['b', 't'].includes(val(a.axPos))) ?? valAxes[0];
    xAxis = axisOf(horizontalAx, 'cat');
    yAxis = axisOf(valAxes.find((a: any) => a !== horizontalAx), 'val');
  } else {
    xAxis = axisOf(catAxes[0], 'cat');
    yAxis = axisOf(valAxes[0], 'val');
  }
  if (xAxis) out.xAxis = xAxis;
  if (yAxis) out.yAxis = yAxis;
  return out;
}
