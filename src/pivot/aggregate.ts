import { PivotSubtotal } from '../types.js';
import { PivotError, PivotValue, isPivotError } from './source.js';

/**
 * Excel's data-field summary functions. Returns null where the pivot shows an
 * empty cell (no applicable values, or #DIV/0! cases), and the first error
 * value when an error takes part in a numeric summary (errors propagate).
 */
export function aggregate(values: PivotValue[], fn: PivotSubtotal = 'sum'): number | null | PivotError {
  if (fn === 'count') {
    const n = values.filter((v) => v !== null).length;
    return n === 0 ? null : n;
  }
  const nums: number[] = [];
  for (const v of values) if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
  if (fn === 'countNums') return nums.length === 0 ? null : nums.length;
  const err = values.find(isPivotError);
  if (err) return err;
  const n = nums.length;
  if (n === 0) return null;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  for (const x of nums) {
    sum += x;
    if (x < min) min = x;
    if (x > max) max = x;
  }
  switch (fn) {
    case 'sum':
      return sum;
    case 'average':
      return sum / n;
    case 'max':
      return max;
    case 'min':
      return min;
    case 'product':
      return nums.reduce((a, b) => a * b, 1);
    case 'var':
    case 'stdDev': {
      if (n < 2) return null;
      const mean = sum / n;
      const v = nums.reduce((a, x) => a + (x - mean) ** 2, 0) / (n - 1);
      return fn === 'var' ? v : Math.sqrt(v);
    }
    case 'varp':
    case 'stdDevp': {
      const mean = sum / n;
      const v = nums.reduce((a, x) => a + (x - mean) ** 2, 0) / n;
      return fn === 'varp' ? v : Math.sqrt(v);
    }
    default:
      return sum;
  }
}

/** Excel's default caption for a summary function ("Sum of Sales"). */
export function defaultDataFieldName(fn: PivotSubtotal, field: string): string {
  const label: Record<PivotSubtotal, string> = {
    sum: 'Sum',
    count: 'Count',
    average: 'Average',
    max: 'Max',
    min: 'Min',
    product: 'Product',
    countNums: 'Count',
    stdDev: 'StdDev',
    stdDevp: 'StdDevp',
    var: 'Var',
    varp: 'Varp',
  };
  return `${label[fn] ?? 'Sum'} of ${field}`;
}
