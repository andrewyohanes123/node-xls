import { unzipSync, strFromU8, unzip } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import {
  WorkBook,
  WorkSheet,
  ParsingOptions,
  CellObject,
  Range,
  ColProperties,
  RowProperties,
} from '../types.js';
import { decode_range, encode_cell, decode_cell } from '../utils/address.js';
import { parseStylesXml } from '../styles/parser.js';
import { parsePivotTableXml } from '../pivot/parser.js';

export function parseSharedStrings(xml: string): string[] {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name) => name === 'si' || name === 'r',
  });

  const parsed = parser.parse(xml);
  const siNodes = parsed.sst?.si || [];

  return siNodes.map((si: any) => {
    if (si.t !== undefined) {
      if (typeof si.t === 'object') {
        return si.t['#text'] !== undefined ? String(si.t['#text']) : '';
      }
      return String(si.t);
    }
    // Rich text runs <r><t>...</t></r>
    if (si.r && Array.isArray(si.r)) {
      return si.r
        .map((rNode: any) => {
          if (rNode.t !== undefined) {
            if (typeof rNode.t === 'object') {
              return rNode.t['#text'] !== undefined ? String(rNode.t['#text']) : '';
            }
            return String(rNode.t);
          }
          return '';
        })
        .join('');
    }
    return '';
  });
}

export function parseXlsxFromFiles(files: Record<string, Uint8Array>, opts?: ParsingOptions): WorkBook {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    isArray: (name) =>
      [
        'sheet',
        'Relationship',
        'col',
        'row',
        'c',
        'mergeCell',
        'pivotCache',
        'pivotTable',
      ].includes(name),
  });

  // 1. Parse styles
  const stylesBytes = files['xl/styles.xml'];
  const styles = stylesBytes ? parseStylesXml(strFromU8(stylesBytes)) : [];

  // 2. Parse shared strings
  const sstBytes = files['xl/sharedStrings.xml'];
  const sharedStrings: string[] = sstBytes
    ? parseSharedStrings(strFromU8(sstBytes))
    : [];

  // 3. Parse workbook relationships
  const wbRelsBytes = files['xl/_rels/workbook.xml.rels'];
  const relMap = new Map<string, string>(); // rId -> target
  if (wbRelsBytes) {
    const wbRelsParsed = parser.parse(strFromU8(wbRelsBytes));
    const relationships = wbRelsParsed.Relationships?.Relationship || [];
    for (const rel of relationships) {
      relMap.set(rel['@_Id'], rel['@_Target']);
    }
  }

  // 4. Parse workbook.xml
  const wbBytes = files['xl/workbook.xml'];
  if (!wbBytes) {
    throw new Error('Invalid XLSX: missing xl/workbook.xml');
  }

  const wbParsed = parser.parse(strFromU8(wbBytes));
  const sheetNodes = wbParsed.workbook?.sheets?.sheet || [];

  const wb: WorkBook = {
    SheetNames: [],
    Sheets: {},
    PivotTables: [],
  };

  const sheetFilePaths: Array<{ name: string; path: string; relId: string }> = [];

  for (const sn of sheetNodes) {
    const name = sn['@_name'];
    const relId = sn['@_r:id'] || sn['@_id'];
    const target = relMap.get(relId);
    if (target) {
      const cleanPath = target.startsWith('xl/') ? target : `xl/${target.replace(/^\//, '')}`;
      sheetFilePaths.push({ name, path: cleanPath, relId });
      wb.SheetNames.push(name);
    }
  }

  // 5. Parse each worksheet
  for (const { name: sheetName, path: sheetPath } of sheetFilePaths) {
    const wsBytes = files[sheetPath];
    if (!wsBytes) continue;

    const wsXml = strFromU8(wsBytes);
    const parsedWs = parser.parse(wsXml);
    const wsNode = parsedWs.worksheet || {};

    const ws: WorkSheet = {};

    // Dimension
    if (wsNode.dimension?.['@_ref']) {
      ws['!ref'] = wsNode.dimension['@_ref'];
    }

    // Cols
    const colNodes = wsNode.cols?.col || [];
    if (colNodes.length > 0) {
      const cols: ColProperties[] = [];
      for (const col of colNodes) {
        const min = Number(col['@_min'] || 1);
        const max = Number(col['@_max'] || min);
        const width = col['@_width'] ? Number(col['@_width']) : undefined;
        const hidden = col['@_hidden'] === '1' || col['@_hidden'] === 'true';
        for (let c = min; c <= max; c++) {
          cols[c - 1] = { width, hidden };
        }
      }
      ws['!cols'] = cols;
    }

    // Rows and cells
    const rowNodes = wsNode.sheetData?.row || [];
    const rows: RowProperties[] = [];
    let minC = Infinity,
      minR = Infinity,
      maxC = -1,
      maxR = -1;

    for (const row of rowNodes) {
      const rIdx = Number(row['@_r']) - 1;
      if (!isNaN(rIdx)) {
        const ht = row['@_ht'] ? Number(row['@_ht']) : undefined;
        const hidden = row['@_hidden'] === '1' || row['@_hidden'] === 'true';
        if (ht !== undefined || hidden) {
          rows[rIdx] = { hpt: ht, hidden };
        }
      }

      const cellNodes = row.c || [];
      for (const cNode of cellNodes) {
        const cellRef = cNode['@_r'];
        if (!cellRef) continue;

        const addr = decode_cell(cellRef);
        minC = Math.min(minC, addr.c);
        minR = Math.min(minR, addr.r);
        maxC = Math.max(maxC, addr.c);
        maxR = Math.max(maxR, addr.r);

        const type = cNode['@_t'];
        const styleIdx = cNode['@_s'] ? Number(cNode['@_s']) : undefined;
        const style = styleIdx !== undefined && styles[styleIdx] ? styles[styleIdx] : undefined;

        let val: any = undefined;
        let cellType = type;

        const vNode = cNode.v;
        const rawVal =
          vNode !== undefined && vNode !== null
            ? typeof vNode === 'object'
              ? vNode['#text']
              : vNode
            : undefined;

        if (type === 's') {
          // Shared string
          const sIdx = Number(rawVal);
          val = !isNaN(sIdx) && sharedStrings[sIdx] !== undefined ? sharedStrings[sIdx] : '';
        } else if (type === 'inlineStr') {
          // Inline string <is><t>...</t></is>
          const isNode = cNode.is;
          if (isNode) {
            val = isNode.t ? (typeof isNode.t === 'object' ? isNode.t['#text'] : isNode.t) : '';
          }
          cellType = 's';
        } else if (type === 'b') {
          val = rawVal === '1' || rawVal === 1 || rawVal === 'true';
        } else if (type === 'e') {
          val = rawVal;
        } else if (type === 'd') {
          val = rawVal ? new Date(rawVal) : undefined;
        } else {
          // Number or formula result
          if (rawVal !== undefined) {
            const numVal = Number(rawVal);
            if (!isNaN(numVal) && String(numVal) === String(rawVal).trim()) {
              val = numVal;
              cellType = 'n';
            } else {
              val = rawVal;
              cellType = 's';
            }
          }
        }

        const cellObj: CellObject = {
          v: val,
          t: cellType,
        };

        if (cNode.f) {
          cellObj.f = typeof cNode.f === 'object' ? cNode.f['#text'] : String(cNode.f);
        }

        if (style) {
          cellObj.s = style;
        }

        ws[cellRef] = cellObj;
      }
    }

    if (rows.length > 0) {
      ws['!rows'] = rows;
    }

    if (!ws['!ref'] && maxC >= 0 && maxR >= 0) {
      ws['!ref'] = `${encode_cell({ c: minC === Infinity ? 0 : minC, r: minR === Infinity ? 0 : minR })}:${encode_cell({
        c: maxC,
        r: maxR,
      })}`;
    }

    // Merges
    const mergeNodes = wsNode.mergeCells?.mergeCell || [];
    if (mergeNodes.length > 0) {
      ws['!merges'] = mergeNodes.map((m: any) => decode_range(m['@_ref']));
    }

    // AutoFilter
    if (wsNode.autoFilter?.['@_ref']) {
      ws['!autofilter'] = { ref: wsNode.autoFilter['@_ref'] };
    }

    // Check sheet rels for pivot tables
    const sheetRelsPath = sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
    const sheetRelsBytes = files[sheetRelsPath];
    if (sheetRelsBytes) {
      const sheetRelsXml = strFromU8(sheetRelsBytes);
      const sheetRelsParsed = parser.parse(sheetRelsXml);
      const sRelationships = sheetRelsParsed.Relationships?.Relationship || [];

      for (const sRel of sRelationships) {
        if (sRel['@_Type']?.includes('pivotTable')) {
          const pivotTarget = sRel['@_Target'];
          const pivotPath = pivotTarget.startsWith('../')
            ? `xl/${pivotTarget.substring(3)}`
            : `xl/worksheets/${pivotTarget}`;

          const pivotBytes = files[pivotPath];
          if (pivotBytes) {
            const pivotXml = strFromU8(pivotBytes);

            // Find pivot table rels to find pivotCacheDefinition
            const pivotRelsPath = pivotPath.replace('pivotTables/', 'pivotTables/_rels/') + '.rels';
            const pivotRelsBytes = files[pivotRelsPath];
            let cacheDefXml = '';

            if (pivotRelsBytes) {
              const pivotRelsParsed = parser.parse(strFromU8(pivotRelsBytes));
              const pRels = pivotRelsParsed.Relationships?.Relationship || [];
              const cacheRel = pRels.find((r: any) => r['@_Type']?.includes('pivotCacheDefinition'));
              if (cacheRel) {
                const cacheTarget = cacheRel['@_Target'];
                const cachePath = cacheTarget.startsWith('../')
                  ? `xl/${cacheTarget.substring(3)}`
                  : `xl/pivotTables/${cacheTarget}`;
                const cacheBytes = files[cachePath];
                if (cacheBytes) {
                  cacheDefXml = strFromU8(cacheBytes);
                }
              }
            }

            if (cacheDefXml) {
              const pivotDef = parsePivotTableXml(pivotXml, cacheDefXml, sheetName);
              if (!ws['!pivots']) {
                ws['!pivots'] = [];
              }
              ws['!pivots'].push(pivotDef);
              wb.PivotTables?.push(pivotDef);
            }
          }
        }
      }
    }

    wb.Sheets[sheetName] = ws;
  }

  return wb;
}

export function readXlsxSync(data: Uint8Array, opts?: ParsingOptions): WorkBook {
  const files = unzipSync(data);
  return parseXlsxFromFiles(files, opts);
}

export async function readXlsxAsync(data: Uint8Array, opts?: ParsingOptions): Promise<WorkBook> {
  return new Promise((resolve, reject) => {
    unzip(data, (err, files) => {
      if (err) reject(err);
      else {
        try {
          resolve(parseXlsxFromFiles(files, opts));
        } catch (e) {
          reject(e);
        }
      }
    });
  });
}
