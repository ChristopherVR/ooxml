import { describe, expect, it } from 'vitest';
import { parseHtmlTable } from './clipboard-html-parse';
import { styleAt } from '../styles';
import { createWorkbook } from '../workbook';

const base = styleAt(createWorkbook(), 0);

const table = (cells: string): string => `<table><tr>${cells}</tr></table>`;

describe('pasted table span limits', () => {
	it('clamps a huge rowspan to the sheet and keeps parsing quickly', () => {
		const started = Date.now();
		const parsed = parseHtmlTable(table('<td rowspan="2000000000">x</td>'), base);
		expect(Date.now() - started).toBeLessThan(2000);
		expect(parsed === undefined || parsed.rows <= 1_048_576).toBe(true);
	});

	it('drops spans that would exceed the cell budget but keeps the cell', () => {
		const parsed = parseHtmlTable(
			table('<td colspan="16384" rowspan="1000">x</td><td>y</td>'),
			base,
		);
		expect(parsed?.data[0]?.[0]?.text).toBe('x');
		expect(parsed?.merges).toEqual([]);
	});

	it('still merges ordinary spans', () => {
		const parsed = parseHtmlTable(table('<td colspan="2" rowspan="2">x</td>'), base);
		expect(parsed?.merges).toHaveLength(1);
		expect(parsed?.rows).toBe(2);
	});
});
