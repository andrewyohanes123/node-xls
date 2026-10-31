/**
 * Excel date serials.
 *
 * Like SheetJS, JS Dates are interpreted by their LOCAL calendar fields by
 * default: `new Date(2026, 0, 15)` is serial 46037 in every timezone.
 * Excel's 1900 system pretends 1900-02-29 exists (serial 60), so serials
 * before 1900-03-01 are shifted by one day.
 */
const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const OFFSET_1904 = 1462;

export interface DateOptions {
  date1904?: boolean;
  /** Use the Date's UTC fields instead of local fields. */
  utc?: boolean;
}

export function dateToSerial(date: Date, opts: DateOptions = {}): number {
  const ms = opts.utc
    ? date.getTime()
    : Date.UTC(
        date.getFullYear(),
        date.getMonth(),
        date.getDate(),
        date.getHours(),
        date.getMinutes(),
        date.getSeconds(),
        date.getMilliseconds()
      );
  const serial = (ms - EPOCH_1900) / DAY_MS;
  if (opts.date1904) return serial - OFFSET_1904;
  return serial < 61 ? serial - 1 : serial;
}

export function serialToDate(serial: number, opts: DateOptions = {}): Date {
  let s = serial;
  if (opts.date1904) s += OFFSET_1904;
  else if (s < 61) s += 1;
  const ms = Math.round(s * DAY_MS) + EPOCH_1900;
  if (opts.utc) return new Date(ms);
  const u = new Date(ms);
  return new Date(
    u.getUTCFullYear(),
    u.getUTCMonth(),
    u.getUTCDate(),
    u.getUTCHours(),
    u.getUTCMinutes(),
    u.getUTCSeconds(),
    u.getUTCMilliseconds()
  );
}

export function isValidDate(v: unknown): v is Date {
  return v instanceof Date && !isNaN(v.getTime());
}

/**
 * Serial for an ISO 8601 date/time string, using its literal calendar fields
 * (any timezone suffix is ignored, as Excel does for t="d" cells).
 */
export function isoToSerial(s: string, opts: DateOptions = {}): number | undefined {
  // time-only values ("13:45:00.000", as SheetJS writes times with cellDates) are day fractions
  const t = /^T?(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?(?:Z|[+-]\d{2}:?\d{2})?$/.exec(s.trim());
  if (t) return (+t[1] * 3600 + +t[2] * 60 + +(t[3] ?? 0) + Number(`0.${t[4] ?? 0}`)) / 86400;
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/.exec(s.trim());
  if (!m) return undefined;
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0), Math.round(Number(`0.${m[7] ?? 0}`) * 1000));
  if (!Number.isFinite(ms)) return undefined;
  return dateToSerial(new Date(ms), { ...opts, utc: true });
}
