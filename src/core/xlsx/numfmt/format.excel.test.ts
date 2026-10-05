import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { formatValue } from './index.js';

/**
 * Every row was recorded from Excel 16 (en-US) by `__fixtures__/generate-excel-cases.ps1`:
 * [value, format code, Range.Text, date1904?].
 */
type Row = [number | string | boolean, string, string, boolean?];

const rows = JSON.parse(
	readFileSync(path.join(import.meta.dirname, '__fixtures__/excel-format-cases.json'), 'utf8'),
) as Row[];

describe('formatValue matches Excel display text', () => {
	it('has the recorded table', () => {
		expect(rows.length).toBeGreaterThan(700);
	});

	it.each(rows)('%j with %s -> %j', (value, format, expected, date1904) => {
		expect(formatValue(value, format, { date1904: date1904 ?? false }).text).toBe(expected);
	});
});
