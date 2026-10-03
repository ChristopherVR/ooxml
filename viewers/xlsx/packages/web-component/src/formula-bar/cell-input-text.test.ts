import { describe, expect, it } from 'vitest';
import { createWorkbook, internStyle, putCell, styleAt, cellError } from '@christophervr/xlsx-core';
import { cellInputText } from '@christophervr/xlsx-core';

const setup = () => {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	const withFormat = (numFmt: string) => internStyle(wb, { ...styleAt(wb, 0), numFmt });
	return { wb, sheet, withFormat };
};

describe('cellInputText', () => {
	it('shows formulas with their equals sign', () => {
		const { wb, sheet } = setup();
		putCell(sheet, 0, 0, { value: 3, formula: 'SUM(B1:B2)' });
		expect(cellInputText(wb, 0, 0, 0)).toBe('=SUM(B1:B2)');
	});

	it('shows plain values, booleans and errors', () => {
		const { wb, sheet } = setup();
		putCell(sheet, 0, 0, { value: 0.1 + 0.2 });
		putCell(sheet, 0, 1, { value: true });
		putCell(sheet, 0, 2, { value: cellError('#N/A') });
		putCell(sheet, 0, 3, { value: 'hello' });
		expect(cellInputText(wb, 0, 0, 0)).toBe('0.3');
		expect(cellInputText(wb, 0, 0, 1)).toBe('TRUE');
		expect(cellInputText(wb, 0, 0, 2)).toBe('#N/A');
		expect(cellInputText(wb, 0, 0, 3)).toBe('hello');
		expect(cellInputText(wb, 0, 5, 5)).toBe('');
	});

	it('keeps text that looks like a number or formula re-typeable with an apostrophe', () => {
		const { wb, sheet } = setup();
		putCell(sheet, 0, 0, { value: '123' });
		putCell(sheet, 0, 1, { value: '=A1' });
		expect(cellInputText(wb, 0, 0, 0)).toBe("'123");
		expect(cellInputText(wb, 0, 0, 1)).toBe("'=A1");
	});

	it('shows dates and percentages the way they are typed', () => {
		const { wb, sheet, withFormat } = setup();
		putCell(sheet, 0, 0, { value: 45292, styleId: withFormat('yyyy-mm-dd') });
		putCell(sheet, 0, 1, { value: 0.15, styleId: withFormat('0%') });
		expect(cellInputText(wb, 0, 0, 0)).toBe('1/1/2024');
		expect(cellInputText(wb, 0, 0, 1)).toBe('15%');
	});

	it('joins rich text runs', () => {
		const { wb, sheet } = setup();
		putCell(sheet, 0, 0, {
			value: 'ab',
			richText: [{ text: 'a' }, { text: 'b', font: { bold: true } }],
		});
		expect(cellInputText(wb, 0, 0, 0)).toBe('ab');
	});
});
