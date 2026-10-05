import { describe, expect, it } from 'vitest';
import type { Table } from '../model.js';
import { structuredToA1 } from './structured-to-a1.js';

const table: Table = {
	id: 1,
	name: 'Sales',
	displayName: 'Sales',
	range: { start: { row: 0, col: 1 }, end: { row: 5, col: 2 } },
	headerRow: true,
	totalsRow: true,
	columns: [{ name: 'Item' }, { name: 'Amount' }],
};
const at = (formulaSheet: string, extra: object = {}) => ({
	table,
	tableSheet: 'Data',
	formulaSheet,
	...extra,
});

describe('structuredToA1', () => {
	it('rewrites columns, specials and bare names', () => {
		expect(structuredToA1('SUM(Sales[Amount])', at('Data'))).toBe('SUM($C$2:$C$5)');
		expect(structuredToA1('Sales[[#Totals],[Amount]]', at('Data'))).toBe('$C$6');
		expect(structuredToA1('Sales[#All]', at('Data'))).toBe('$B$1:$C$6');
		expect(structuredToA1('Sales[#Headers]', at('Data'))).toBe('$B$1:$C$1');
		expect(structuredToA1('ROWS(Sales)', at('Other'))).toBe('ROWS(Data!$B$2:$C$5)');
		expect(structuredToA1('[@Amount]*2', at('Data', { row: 3, inside: true }))).toBe('$C4*2');
		expect(structuredToA1('Sales[Missing]', at('Data'))).toBe('#REF!');
	});
	it('leaves other tables and strings alone', () => {
		expect(structuredToA1('Other[A]&"Sales[A]"', at('Data'))).toBe('Other[A]&"Sales[A]"');
	});
});
