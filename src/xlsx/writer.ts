import { zipSync, strToU8, zip } from 'fflate';
import {
  WorkBook,
  WorkSheet,
  WritingOptions,
  CellObject,
  PivotTableDefinition,
  OutputType,
} from '../types.js';
import {
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
} from '../utils/address.js';
import { StyleBuilder, escapeXml } from '../styles/index.js';
import { generateOpenXmlPivotParts } from '../pivot/openxml.js';

export class SharedStringsBuilder {
  private strings: string[] = [];
  private stringMap = new Map<string, number>();

  public getIndex(str: string): number {
    if (this.stringMap.has(str)) {
      return this.stringMap.get(str)!;
    }
    const idx = this.strings.length;
    this.strings.push(str);
    this.stringMap.set(str, idx);
    return idx;
  }

  public get count(): number {
    return this.strings.length;
  }

  public buildXml(): string {
    const itemsXml = this.strings
      .map((s) => {
        const needsPreserve = /^\s|\s$|\n|\r|\t/.test(s);
        const spaceAttr = needsPreserve ? ' xml:space="preserve"' : '';
        return `<si><t${spaceAttr}>${escapeXml(s)}</t></si>`;
      })
      .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.strings.length}" uniqueCount="${this.strings.length}">${itemsXml}</sst>`;
  }
}

export function buildXlsxArchiveFiles(wb: WorkBook, opts?: WritingOptions): Record<string, Uint8Array> {
  const files: Record<string, Uint8Array> = {};
  const styleBuilder = new StyleBuilder();
  const sst = new SharedStringsBuilder();

  const sheetNames = wb.SheetNames && wb.SheetNames.length > 0 ? wb.SheetNames : ['Sheet1'];
  if (!wb.Sheets[sheetNames[0]]) {
    wb.Sheets[sheetNames[0]] = {};
  }

  // Collect pivot tables
  const allPivotTables: PivotTableDefinition[] = [];
  if (wb.PivotTables) {
    allPivotTables.push(...wb.PivotTables);
  }
  for (const sName of sheetNames) {
    const ws = wb.Sheets[sName];
    if (ws && ws['!pivots']) {
      for (const p of ws['!pivots']) {
        if (!allPivotTables.some((ep) => ep.name === p.name)) {
          allPivotTables.push(p);
        }
      }
    }
  }

  // Ensure pivot IDs
  allPivotTables.forEach((p, idx) => {
    p.id = idx + 1;
  });

  // Build Worksheets XML
  sheetNames.forEach((sheetName, index) => {
    const ws = wb.Sheets[sheetName] || {};
    const sheetNum = index + 1;

    // Cells
    const cellKeys = Object.keys(ws).filter((k) => !k.startsWith('!'));
    const rowsMap = new Map<number, Array<{ c: number; key: string; cell: CellObject }>>();

    let minC = Infinity,
      minR = Infinity,
      maxC = -1,
      maxR = -1;

    for (const key of cellKeys) {
      try {
        const addr = decode_cell(key);
        minC = Math.min(minC, addr.c);
        minR = Math.min(minR, addr.r);
        maxC = Math.max(maxC, addr.c);
        maxR = Math.max(maxR, addr.r);

        let rowArr = rowsMap.get(addr.r);
        if (!rowArr) {
          rowArr = [];
          rowsMap.set(addr.r, rowArr);
        }
        rowArr.push({ c: addr.c, key, cell: ws[key] });
      } catch {
        // ignore invalid cell refs
      }
    }

    if (ws['!ref']) {
      try {
        const refRange = decode_range(ws['!ref']);
        minC = Math.min(minC, refRange.s.c);
        minR = Math.min(minR, refRange.s.r);
        maxC = Math.max(maxC, refRange.e.c);
        maxR = Math.max(maxR, refRange.e.r);
      } catch {
        // ignore
      }
    }

    const hasCells = maxC >= 0 && maxR >= 0;
    const dimensionRef = hasCells
      ? `${encode_cell({ c: minC === Infinity ? 0 : minC, r: minR === Infinity ? 0 : minR })}:${encode_cell({
          c: maxC,
          r: maxR,
        })}`
      : 'A1';

    // Cols
    let colsXml = '';
    if (ws['!cols'] && ws['!cols'].length > 0) {
      const colItems = ws['!cols']
        .map((col, cIdx) => {
          let width = 10;
          let customWidth = false;
          if (col.width !== undefined) {
            width = col.width;
            customWidth = true;
          } else if (col.wch !== undefined) {
            width = col.wch + 1;
            customWidth = true;
          } else if (col.wpx !== undefined) {
            width = col.wpx / 7;
            customWidth = true;
          }
          const colNum = cIdx + 1;
          const hidden = col.hidden ? ' hidden="1"' : '';
          const cwAttr = customWidth ? ' customWidth="1"' : '';
          return `<col min="${colNum}" max="${colNum}" width="${width}"${cwAttr}${hidden}/>`;
        })
        .join('');
      colsXml = `<cols>${colItems}</cols>`;
    }

    // Rows & Cells XML
    const sortedRowIndices = Array.from(rowsMap.keys()).sort((a, b) => a - b);
    const rowXmls: string[] = [];

    for (const r of sortedRowIndices) {
      const cells = rowsMap.get(r)!;
      cells.sort((a, b) => a.c - b.c);

      const rowProp = ws['!rows']?.[r];
      let rowAttrs = `r="${r + 1}"`;
      if (rowProp) {
        if (rowProp.hpt !== undefined) {
          rowAttrs += ` ht="${rowProp.hpt}" customHeight="1"`;
        } else if (rowProp.hpx !== undefined) {
          rowAttrs += ` ht="${rowProp.hpx * 0.75}" customHeight="1"`;
        }
        if (rowProp.hidden) {
          rowAttrs += ' hidden="1"';
        }
      }

      const cellXmls = cells.map(({ key, cell }) => {
        let styleIdx = 0;
        if (cell.s) {
          styleIdx = styleBuilder.getStyleIndex(cell.s);
        }

        const sAttr = styleIdx > 0 ? ` s="${styleIdx}"` : '';

        // Value handling
        let typeAttr = '';
        let valXml = '';
        const formulaXml = cell.f ? `<f>${escapeXml(cell.f)}</f>` : '';

        if (cell.t === 'b') {
          typeAttr = ' t="b"';
          valXml = `<v>${cell.v ? '1' : '0'}</v>`;
        } else if (cell.t === 'n') {
          typeAttr = '';
          valXml = cell.v !== undefined && cell.v !== null ? `<v>${cell.v}</v>` : '';
        } else if (cell.t === 'e') {
          typeAttr = ' t="e"';
          valXml = `<v>${escapeXml(String(cell.v || ''))}</v>`;
        } else if (cell.t === 'd' || cell.v instanceof Date) {
          typeAttr = ' t="d"';
          valXml = `<v>${(cell.v instanceof Date ? cell.v : new Date(cell.v)).toISOString()}</v>`;
        } else if (cell.t === 's' || typeof cell.v === 'string') {
          const strVal = String(cell.v ?? '');
          const sstIdx = sst.getIndex(strVal);
          typeAttr = ' t="s"';
          valXml = `<v>${sstIdx}</v>`;
        } else if (cell.v !== undefined && cell.v !== null) {
          if (typeof cell.v === 'number') {
            valXml = `<v>${cell.v}</v>`;
          } else {
            const strVal = String(cell.v);
            const sstIdx = sst.getIndex(strVal);
            typeAttr = ' t="s"';
            valXml = `<v>${sstIdx}</v>`;
          }
        }

        return `<c r="${key}"${typeAttr}${sAttr}>${formulaXml}${valXml}</c>`;
      });

      rowXmls.push(`<row ${rowAttrs}>${cellXmls.join('')}</row>`);
    }

    const sheetDataXml = `<sheetData>${rowXmls.join('')}</sheetData>`;

    // Merge Cells
    let mergesXml = '';
    if (ws['!merges'] && ws['!merges'].length > 0) {
      const mergeItems = ws['!merges'].map((m) => `<mergeCell ref="${encode_range(m)}"/>`).join('');
      mergesXml = `<mergeCells count="${ws['!merges'].length}">${mergeItems}</mergeCells>`;
    }

    // AutoFilter
    let autoFilterXml = '';
    if (ws['!autofilter']?.ref) {
      autoFilterXml = `<autoFilter ref="${escapeXml(ws['!autofilter'].ref)}"/>`;
    }

    const fullSheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="${dimensionRef}"/><sheetViews><sheetView workbookViewId="0"${
      index === 0 ? ' tabSelected="1"' : ''
    }/></sheetViews><sheetFormatPr defaultRowHeight="15"/>${colsXml}${sheetDataXml}${mergesXml}${autoFilterXml}</worksheet>`;

    files[`xl/worksheets/sheet${sheetNum}.xml`] = strToU8(fullSheetXml);

    // Pivot tables on this sheet
    const sheetPivots = allPivotTables.filter((p) => p.targetSheet === sheetName);
    if (sheetPivots.length > 0) {
      const rels = sheetPivots.map((p) => {
        return `<Relationship Id="rIdPivot${p.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable${p.id}.xml"/>`;
      });
      const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join(
        ''
      )}</Relationships>`;
      files[`xl/worksheets/_rels/sheet${sheetNum}.xml.rels`] = strToU8(sheetRelsXml);
    }
  });

  // Generate Pivot XMLs
  for (const pivot of allPivotTables) {
    const srcSheet = wb.Sheets[pivot.sourceSheet];
    if (srcSheet) {
      const parts = generateOpenXmlPivotParts(srcSheet, pivot, pivot.id!);
      files[`xl/pivotCache/pivotCacheDefinition${pivot.id}.xml`] = strToU8(parts.cacheDefinitionXml);
      files[`xl/pivotCache/pivotCacheRecords${pivot.id}.xml`] = strToU8(parts.cacheRecordsXml);
      files[`xl/pivotCache/_rels/pivotCacheDefinition${pivot.id}.xml.rels`] = strToU8(parts.cacheDefRelsXml);
      files[`xl/pivotTables/pivotTable${pivot.id}.xml`] = strToU8(parts.pivotTableXml);
      files[`xl/pivotTables/_rels/pivotTable${pivot.id}.xml.rels`] = strToU8(parts.pivotTableRelsXml);
    }
  }

  // Shared Strings
  if (sst.count > 0) {
    files['xl/sharedStrings.xml'] = strToU8(sst.buildXml());
  }

  // Styles
  files['xl/styles.xml'] = strToU8(styleBuilder.buildXml());

  // xl/workbook.xml
  const sheetsXml = sheetNames
    .map((name, idx) => `<sheet name="${escapeXml(name)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`)
    .join('');

  let pivotCachesXml = '';
  if (allPivotTables.length > 0) {
    const cacheNodes = allPivotTables
      .map((p) => `<pivotCache cacheId="${p.id}" r:id="rIdCache${p.id}"/>`)
      .join('');
    pivotCachesXml = `<pivotCaches>${cacheNodes}</pivotCaches>`;
  }

  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheetsXml}</sheets>${pivotCachesXml}</workbook>`;
  files['xl/workbook.xml'] = strToU8(workbookXml);

  // xl/_rels/workbook.xml.rels
  const wbRels: string[] = [];
  sheetNames.forEach((_, idx) => {
    wbRels.push(
      `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`
    );
  });

  const nextRelId = sheetNames.length + 1;
  wbRels.push(
    `<Relationship Id="rId${nextRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );

  let curRelId = nextRelId + 1;
  if (sst.count > 0) {
    wbRels.push(
      `<Relationship Id="rId${curRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`
    );
    curRelId++;
  }

  for (const pivot of allPivotTables) {
    wbRels.push(
      `<Relationship Id="rIdCache${pivot.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${pivot.id}.xml"/>`
    );
  }

  const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels.join(
    ''
  )}</Relationships>`;
  files['xl/_rels/workbook.xml.rels'] = strToU8(workbookRelsXml);

  // _rels/.rels
  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  files['_rels/.rels'] = strToU8(rootRelsXml);

  // docProps/core.xml
  const coreProps = wb.Props || {};
  const creator = escapeXml(coreProps.Author || 'node-xls');
  const title = escapeXml(coreProps.Title || '');
  const corePropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>${creator}</dc:creator><dc:title>${title}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created></cp:coreProperties>`;
  files['docProps/core.xml'] = strToU8(corePropsXml);

  // docProps/app.xml
  const appPropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>node-xls</Application></Properties>`;
  files['docProps/app.xml'] = strToU8(appPropsXml);

  // [Content_Types].xml
  const ctOverrides: string[] = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
  ];

  if (sst.count > 0) {
    ctOverrides.push(
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    );
  }

  sheetNames.forEach((_, idx) => {
    ctOverrides.push(
      `<Override PartName="/xl/worksheets/sheet${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  });

  for (const pivot of allPivotTables) {
    ctOverrides.push(
      `<Override PartName="/xl/pivotTables/pivotTable${pivot.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`,
      `<Override PartName="/xl/pivotCache/pivotCacheDefinition${pivot.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`,
      `<Override PartName="/xl/pivotCache/pivotCacheRecords${pivot.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`
    );
  }

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${ctOverrides.join(
    ''
  )}</Types>`;
  files['[Content_Types].xml'] = strToU8(contentTypesXml);

  return files;
}

export function writeXlsxSync(wb: WorkBook, opts?: WritingOptions): Uint8Array {
  const files = buildXlsxArchiveFiles(wb, opts);
  return zipSync(files, { level: opts?.compression === false ? 0 : 6 });
}

export async function writeXlsxAsync(wb: WorkBook, opts?: WritingOptions): Promise<Uint8Array> {
  const files = buildXlsxArchiveFiles(wb, opts);
  return new Promise((resolve, reject) => {
    zip(files, { level: opts?.compression === false ? 0 : 6 }, (err, data) => {
      if (err) reject(err);
      else resolve(data);
    });
  });
}

export function convertOutputType(bytes: Uint8Array, type?: OutputType): any {
  const outType = type || (typeof Buffer !== 'undefined' ? 'buffer' : 'array');
  switch (outType) {
    case 'buffer':
      return Buffer.from(bytes);
    case 'array':
      return bytes;
    case 'binary': {
      let binary = '';
      for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      return binary;
    }
    case 'base64': {
      if (typeof Buffer !== 'undefined') {
        return Buffer.from(bytes).toString('base64');
      }
      let binary = '';
      for (let i = 0; i < bytes.length; i++) {
        binary += String.fromCharCode(bytes[i]);
      }
      return btoa(binary);
    }
    case 'string': {
      const decoder = new TextDecoder('utf-8');
      return decoder.decode(bytes);
    }
    default:
      return Buffer.from(bytes);
  }
}
