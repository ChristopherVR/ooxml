// Compares the engine with results recorded from real Excel (see __fixtures__/excel-cases.ps1).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { CellValue } from '../model';
import { isCellError } from '../model';
import { formatValue } from '../numfmt/format';
import { book, engine, get } from './test-helpers';
import { parseAddress } from '../address';

interface ExcelResult {
	formula: string;
	type: 'number' | 'string' | 'boolean' | 'error' | 'empty' | 'rejected';
	value?: number | string | boolean;
	/** What the cell shows in a wide General column (numbers, in the newer corpora). */
	text?: string;
}

const fixture = (name: string): string =>
	readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8');

function dataValue(text: string): { value: CellValue; formula?: string } {
	if (text.startsWith('=')) return { value: null, formula: text.slice(1) };
	if (text.startsWith("'")) return { value: text.slice(1) };
	if (/^[+-]?\d+(\.\d+)?$/.test(text)) return { value: Number(text) };
	if (text === 'TRUE' || text === 'FALSE') return { value: text === 'TRUE' };
	if (text === '') return { value: null };
	return { value: text };
}

function buildWorkbook(lines: readonly string[]) {
	const wb = book({}, ['Sheet1']);
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('no sheet');
	const formulas: string[] = [];
	for (const line of lines) {
		if (line.startsWith('#') || line.trim() === '') continue;
		if (line.startsWith('!')) {
			const eq = line.indexOf('=');
			const at = parseAddress(line.slice(1, eq));
			if (!at) throw new Error(line);
			const { value, formula } = dataValue(line.slice(eq + 1));
			if (value === null && formula === undefined) continue;
			putCell(sheet, at.row, at.col, formula === undefined ? { value } : { value, formula });
			continue;
		}
		formulas.push(line);
	}
	formulas.forEach((f, i) => putCell(sheet, i, 25, { value: null, formula: f.slice(1) }));
	return { wb, formulas };
}

describe.each(['excel', 'excel-p4'])('Excel parity corpus %s', (corpus) => {
	const lines = fixture(`${corpus}-cases.txt`).split(/\r?\n/);
	const results = JSON.parse(fixture(`${corpus}-results.json`)) as ExcelResult[];
	const { wb, formulas } = buildWorkbook(lines);
	engine(wb).recalculateAll();

	it('has a recorded Excel result for every case', () => {
		expect(results.map((r) => r.formula)).toEqual(formulas);
	});

	const cases = results.map((r, i) => ({ ...r, row: i + 1 })).filter((r) => r.type !== 'rejected');
	it.each(cases)('$formula', ({ row, type, value, formula, text }) => {
		const actual = get(wb, `Z${row}`);
		switch (type) {
			case 'number': {
				expect(typeof actual).toBe('number');
				const expected = value as number;
				// Excel stops its IRR-style iterations early; compare those more loosely.
				const relative = /IRR|RATE/.test(formula) ? 1e-7 : 1e-9;
				const tolerance = Math.max(1e-9, Math.abs(expected) * relative);
				expect(Math.abs((actual as number) - expected)).toBeLessThanOrEqual(tolerance);
				if (text !== undefined) expect(formatValue(actual, 'General').text).toBe(text);
				break;
			}
			case 'error':
				expect(isCellError(actual) ? actual.error : actual).toBe(value);
				break;
			case 'empty':
				expect(actual === null || actual === 0 || actual === '').toBe(true);
				break;
			default:
				expect(actual).toBe(value);
		}
	});
});
