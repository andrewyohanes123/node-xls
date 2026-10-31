import { PivotSubtotal } from '../types.js';
import { PivotValue } from './source.js';

/**
 * Excel's data-field summary functions. Returns null where Excel shows an
 * empty cell or #DIV/0! (no applicable values).
 */
export function aggregate(values: PivotValue[], fn: PivotSubtotal = 'sum'): number | null {
  if (fn === 'count') return values.filter((v) => v !== null).length;
  const nums = values.filter((v): v is number => typeof v === 'number' && Number.isFinite(v));
  if (fn === 'countNums') return nums.length;
  if (fn === 'sum') return nums.reduce((a, b) => a + b, 0);
  if (nums.length === 0) return fn === 'max' || fn === 'min' || fn === 'product' ? 0 : null;
  const n = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  switch (fn) {
    case 'average':
      return sum / n;
    case 'max':
      return Math.max(...nums);
    case 'min':
      return Math.min(...nums);
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
