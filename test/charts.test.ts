/**
 * Charts (SheetJS "Pro Chart" tier): every chart type is validated against
 * dml-chart.xsd, its embedded caches are checked, and it is read back.
 */
import { describe, it, expect } from 'vitest';
import XLSX, { utils, SheetChart } from '../src/index.js';
import { book, roundtrip } from './helpers/util.js';
import { makePng } from './helpers/png.js';

function dataBook() {
  const ws = utils.aoa_to_sheet([
    ['Month', 'North', 'South', 'Share'],
    ['Jan', 120, 80, 0.4],
    ['Feb', 135, 95, 0.35],
    ['Mar', 150, 70, 0.25],
  ]);
  ws.B2.z = '#,##0';
  const wb = book(['Sales Data', ws], ['Dash', utils.aoa_to_sheet([['Charts']])]);
  return wb;
}

const series = [
  { name: "'Sales Data'!$B$1", categories: "'Sales Data'!$A$2:$A$4", values: "'Sales Data'!$B$2:$B$4" },
  { name: 'South', categories: "'Sales Data'!A2:A4", values: "'Sales Data'!C2:C4", color: { rgb: 'ED7D31' } },
];

describe('chart types', () => {
  const cases: SheetChart[] = [
    { type: 'column', title: 'Sales by month', series, anchor: { from: 'A3' } },
    { type: 'column', grouping: 'stacked', series, anchor: { from: 'A3' }, dataLabels: true },
    { type: 'bar', grouping: 'percentStacked', series, anchor: { from: 'A3' }, gapWidth: 80 },
    { type: 'line', series: series.map((s) => ({ ...s, marker: true, smooth: true })), anchor: { from: 'A3' }, yAxis: { title: 'Units', min: 0, max: 200, numFmt: '#,##0' }, xAxis: { title: 'Month' } },
    { type: 'area', grouping: 'stacked', series, anchor: { from: 'A3' }, legend: 'b' },
    { type: 'pie', title: 'Share', series: [{ categories: "'Sales Data'!A2:A4", values: "'Sales Data'!D2:D4" }], anchor: { from: 'A3' }, dataLabels: true },
    { type: 'doughnut', holeSize: 60, series: [{ categories: "'Sales Data'!A2:A4", values: "'Sales Data'!D2:D4" }], anchor: { from: 'A3' }, legend: false },
    { type: 'scatter', series: [{ name: 'N vs S', categories: "'Sales Data'!B2:B4", values: "'Sales Data'!C2:C4" }], anchor: { from: 'A3' } },
  ];
  for (const chart of cases) {
    it(`${chart.type}${chart.grouping ? ` (${chart.grouping})` : ''}`, async () => {
      const wb = dataBook();
      wb.Sheets.Dash['!charts'] = [chart];
      const { rep, back } = await roundtrip(wb);
      expect(rep.contentTypes['xl/charts/chart1.xml']).toBe('application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const read = back.Sheets.Dash['!charts']![0];
      expect(read.type).toBe(chart.type);
      expect(read.series.length).toBe(chart.series.length);
      if (chart.grouping && chart.grouping !== 'clustered') expect(read.grouping).toBe(chart.grouping);
      if (chart.title) expect(read.title).toBe(chart.title);
      expect(read.series[0].values).toMatch(/^'Sales Data'!\$[BCD]\$2:\$[BCD]\$4$/);
    });
  }
});

describe('chart data', () => {
  it('embeds value/category/name caches and absolute references', async () => {
    const wb = dataBook();
    wb.Sheets.Dash['!charts'] = [{ type: 'column', series, anchor: { from: 'B2', to: 'J20' } }];
    const { rep } = await roundtrip(wb);
    const xml = rep.xml('xl/charts/chart1.xml');
    expect(xml).toContain("<c:f>'Sales Data'!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>North</c:v></c:pt></c:strCache>");
    expect(xml).toContain('<c:tx><c:v>South</c:v></c:tx>');
    expect(xml).toContain("<c:f>'Sales Data'!$A$2:$A$4</c:f><c:strCache><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>Jan</c:v></c:pt>");
    expect(xml).toContain('<c:numCache><c:formatCode>#,##0</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>120</c:v></c:pt><c:pt idx="1"><c:v>135</c:v></c:pt>');
    expect(xml).toContain('<a:srgbClr val="ED7D31"/>');
    expect(xml).toContain('<a:schemeClr val="accent1"/>');
    const drawing = rep.xml('xl/drawings/drawing1.xml');
    expect(drawing).toContain('<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>');
    expect(drawing).toContain('<xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row>');
  });

  it('missing values become gaps; numeric categories use number caches', async () => {
    const ws = utils.aoa_to_sheet([
      [2024, 10],
      [2025, null],
      [2026, 30],
    ]);
    ws['!charts'] = [{ type: 'line', series: [{ categories: 'A1:A3', values: 'B1:B3' }], anchor: { from: 'D1' } }];
    const { rep } = await roundtrip(book(['S', ws]));
    const xml = rep.xml('xl/charts/chart1.xml');
    expect(xml).toContain('<c:cat><c:numRef><c:f>S!$A$1:$A$3</c:f>');
    expect(xml).toContain('<c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt>');
  });

  it('charts and images share one drawing; read back with options', async () => {
    const wb = dataBook();
    wb.Sheets.Dash['!images'] = [{ data: makePng(10, 10, () => [0, 0, 0, 255]), anchor: { from: 'L1' } }];
    wb.Sheets.Dash['!charts'] = [
      { type: 'line', name: 'Trend', title: 'Trend', series, anchor: { from: 'A3', width: 400, height: 250 }, yAxis: { title: 'Units', min: 0, max: 200 }, legend: 't' },
      { type: 'pie', series: [{ values: "'Sales Data'!D2:D4" }], anchor: { from: 'H3' } },
    ];
    const { rep, back } = await roundtrip(wb);
    expect(Object.keys(rep.files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))).toEqual(['xl/drawings/drawing1.xml']);
    expect(Object.keys(rep.files).filter((f) => f.startsWith('xl/charts/'))).toHaveLength(2);
    const charts = back.Sheets.Dash['!charts']!;
    expect(charts[0]).toMatchObject({ type: 'line', name: 'Trend', title: 'Trend', legend: 't', yAxis: { title: 'Units', min: 0, max: 200 } });
    expect(charts[0].series[1]).toMatchObject({ name: 'South', color: { rgb: 'ED7D31' } });
    expect(back.Sheets.Dash['!images']).toHaveLength(1);
    // and the read-back definitions can be written again
    const again = await roundtrip(back);
    expect(again.back.Sheets.Dash['!charts']![0].series[0].values).toBe("'Sales Data'!$B$2:$B$4");
  });

  it('ExcelJS loads workbooks containing charts', async () => {
    const wb = dataBook();
    wb.Sheets.Dash['!charts'] = [{ type: 'column', series, anchor: { from: 'A3' } }];
    const { rep } = await roundtrip(wb);
    expect(rep.exceljs.getWorksheet('Dash')).toBeDefined();
  });

  it('rejects bad references', () => {
    const write = (chart: any) => () => {
      const wb = dataBook();
      wb.Sheets.Dash['!charts'] = [chart];
      XLSX.write(wb, { type: 'array' });
    };
    expect(write({ type: 'column', series: [{ values: 'Nope!A1:A3' }], anchor: { from: 'A1' } })).toThrow(/sheet "Nope" not found/);
    expect(write({ type: 'column', series: [{ values: 'garbage' }], anchor: { from: 'A1' } })).toThrow(/invalid range/);
    expect(write({ type: 'column', series: [], anchor: { from: 'A1' } })).toThrow(/at least one series/);
    expect(write({ type: 'radar', series, anchor: { from: 'A1' } })).toThrow(/Unsupported chart type/);
  });
});
