/**
 * Sheet and workbook protection with Excel's SHA-512 password verifier
 * (ECMA-376 Part 1, 18.2.29): H0 = SHA512(salt + UTF-16LE(password)),
 * Hn = SHA512(Hn-1 + LE32(n-1)) for spinCount iterations.
 *
 * Note: sheet/workbook protection is an editing restriction, not encryption.
 */
import { createHash, randomBytes } from 'node:crypto';
import { SheetProtection, WorkbookProtection } from '../types.js';
import { attrs, AttrValue } from '../xlsx/xml.js';

export const DEFAULT_SPIN_COUNT = 100_000;

export interface PasswordHash {
  algorithmName: 'SHA-512';
  hashValue: string;
  saltValue: string;
  spinCount: number;
}

export function hashPassword(password: string, saltValue?: string, spinCount = DEFAULT_SPIN_COUNT): PasswordHash {
  const salt = saltValue ? Buffer.from(saltValue, 'base64') : randomBytes(16);
  let key = createHash('sha512').update(salt).update(Buffer.from(password, 'utf16le')).digest();
  const iterator = Buffer.alloc(4);
  for (let i = 0; i < spinCount; i++) {
    iterator.writeUInt32LE(i, 0);
    key = createHash('sha512').update(key).update(iterator).digest();
  }
  return { algorithmName: 'SHA-512', hashValue: key.toString('base64'), saltValue: salt.toString('base64'), spinCount };
}

/** Attributes whose OOXML default is false (action allowed unless set). */
const DEFAULT_FALSE = ['objects', 'scenarios', 'selectLockedCells', 'selectUnlockedCells'] as const;
/** Attributes whose OOXML default is true (action locked unless set false). */
const DEFAULT_TRUE = [
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'sort',
  'autoFilter',
  'pivotTables',
] as const;

export function sheetProtectionXml(p?: SheetProtection): string {
  if (!p) return '';
  const list: Array<[string, AttrValue]> = [];
  if (p.password) {
    const h = hashPassword(p.password);
    list.push(['algorithmName', h.algorithmName], ['hashValue', h.hashValue], ['saltValue', h.saltValue], ['spinCount', h.spinCount]);
  } else if (p.hashValue) {
    list.push(['algorithmName', p.algorithmName], ['hashValue', p.hashValue], ['saltValue', p.saltValue], ['spinCount', p.spinCount]);
  } else if (p.passwordHash) {
    list.push(['password', p.passwordHash]);
  }
  list.push(['sheet', true]);
  for (const k of DEFAULT_FALSE) if (p[k]) list.push([k, '1']);
  for (const k of DEFAULT_TRUE) if (p[k] === false) list.push([k, '0']);
  return `<sheetProtection${attrs(list)}/>`;
}

export function workbookProtectionXml(p?: WorkbookProtection): string {
  if (!p) return '';
  const list: Array<[string, AttrValue]> = [];
  if (p.password) {
    const h = hashPassword(p.password);
    list.push(
      ['workbookAlgorithmName', h.algorithmName],
      ['workbookHashValue', h.hashValue],
      ['workbookSaltValue', h.saltValue],
      ['workbookSpinCount', h.spinCount]
    );
  } else if (p.hashValue) {
    list.push(
      ['workbookAlgorithmName', p.algorithmName],
      ['workbookHashValue', p.hashValue],
      ['workbookSaltValue', p.saltValue],
      ['workbookSpinCount', p.spinCount]
    );
  }
  list.push(['lockStructure', p.lockStructure !== false], ['lockWindows', p.lockWindows === true]);
  return `<workbookProtection${attrs(list)}/>`;
}
