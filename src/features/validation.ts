/**
 * Data validation (<dataValidations>): dropdown lists, numeric/date/length
 * bounds and custom formulas, with input prompts and error alerts.
 */
import { DataValidation } from '../types.js';
import { WriteContext } from '../xlsx/context.js';
import { attrs, escapeXml } from '../xlsx/xml.js';
import { dateToSerial } from '../utils/dates.js';

function formulaOf(v: string | number | Date, ctx: WriteContext): string {
  if (v instanceof Date) return String(dateToSerial(v, { date1904: ctx.date1904 }));
  if (typeof v === 'number') return String(v);
  return v.replace(/^=/, '');
}

function checkLength(label: string, s: string | undefined, max: number) {
  if (s !== undefined && s.length > max) throw new Error(`Data validation ${label} exceeds ${max} characters`);
}

export function dataValidationsXml(list: DataValidation[], ctx: WriteContext): string {
  const items = list.map((dv) => {
    let f1: string | undefined;
    let f2: string | undefined;
    if (dv.type === 'list' && dv.values) {
      const joined = dv.values.map((v) => String(v).replace(/"/g, '""')).join(',');
      if (joined.length > 255) {
        throw new Error('Inline list validation exceeds 255 characters; put the values in cells and use formula1 with a range');
      }
      f1 = `"${joined}"`;
    } else if (dv.formula1 !== undefined) {
      f1 = formulaOf(dv.formula1, ctx);
    }
    if (dv.formula2 !== undefined) f2 = formulaOf(dv.formula2, ctx);
    if (f1 === undefined && dv.type !== 'any') throw new Error(`Data validation of type ${dv.type} needs values or formula1`);
    const needsTwo = !['any', 'list', 'custom'].includes(dv.type) && (dv.operator ?? 'between').includes('etween');
    if (needsTwo && f2 === undefined) throw new Error(`Data validation operator ${dv.operator ?? 'between'} needs formula2`);

    checkLength('prompt title', dv.prompt?.title, 32);
    checkLength('prompt', dv.prompt?.message, 255);
    checkLength('error title', dv.error?.title, 32);
    checkLength('error message', dv.error?.message, 255);

    const a = attrs([
      ['type', dv.type === 'any' ? undefined : dv.type],
      ['errorStyle', dv.error?.style && dv.error.style !== 'stop' ? dv.error.style : undefined],
      ['operator', ['any', 'list', 'custom'].includes(dv.type) || !dv.operator || dv.operator === 'between' ? undefined : dv.operator],
      ['allowBlank', dv.allowBlank !== false],
      ['showDropDown', dv.type === 'list' && dv.hideDropDown === true],
      ['showInputMessage', true],
      ['showErrorMessage', true],
      ['errorTitle', dv.error?.title],
      ['error', dv.error?.message],
      ['promptTitle', dv.prompt?.title],
      ['prompt', dv.prompt?.message],
      ['sqref', dv.ref.replace(/\$/g, '').trim()],
    ]);
    const children = (f1 !== undefined ? `<formula1>${escapeXml(f1)}</formula1>` : '') + (f2 !== undefined ? `<formula2>${escapeXml(f2)}</formula2>` : '');
    return `<dataValidation${a}>${children}</dataValidation>`;
  });
  return `<dataValidations count="${items.length}">${items.join('')}</dataValidations>`;
}
