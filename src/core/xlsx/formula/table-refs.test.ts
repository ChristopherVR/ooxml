import { describe, expect, it } from 'vitest';
import { renameTableColumnInFormula, renameTableInFormula } from './table-refs';

describe('renameTableInFormula', () => {
	it('renames structured references and bare table names', () => {
		expect(renameTableInFormula('SUM(Sales[Amount])+Sales[@Qty]', 'sales', 'Sales_2')).toBe(
			'SUM(Sales_2[Amount])+Sales_2[@Qty]',
		);
		expect(renameTableInFormula('SUBTOTAL(109,Sales[[#Totals],[A]])', 'Sales', 'T2')).toBe(
			'SUBTOTAL(109,T2[[#Totals],[A]])',
		);
		expect(renameTableInFormula('ROWS(Sales)', 'Sales', 'T2')).toBe('ROWS(T2)');
	});
	it('leaves other names, strings and unparsable text alone', () => {
		expect(renameTableInFormula('"Sales[A]"&SalesTax[A]', 'Sales', 'T2')).toBe(
			'"Sales[A]"&SalesTax[A]',
		);
		expect(renameTableInFormula('SUM((', 'Sales', 'T2')).toBe('SUM((');
	});
});

describe('renameTableColumnInFormula', () => {
	const rename = (formula: string, inside = false) =>
		renameTableColumnInFormula(formula, 'Sales', 'Amount', 'Cost', inside);

	it('renames column leaves, row references and range endpoints case-insensitively', () => {
		expect(rename('SUM(sales[amount])+Sales[@Amount]')).toBe('SUM(sales[Cost])+Sales[@Cost]');
		expect(rename('Sales[[#Totals],[Amount]]+Sales[[Amount]:[Qty]]')).toBe(
			'Sales[[#Totals],[Cost]]+Sales[[Cost]:[Qty]]',
		);
		expect(rename('Sales[@[Amount]]')).toBe('Sales[@[Cost]]');
	});

	it('limits bare references to formulas inside the renamed table', () => {
		expect(rename('[@Amount]+SUM([Amount])', true)).toBe('[@Cost]+SUM([Cost])');
		expect(rename('[@Amount]+Other[Amount]+"Sales[Amount]"')).toBe(
			'[@Amount]+Other[Amount]+"Sales[Amount]"',
		);
		expect(rename("'Book.xlsx'!Sales[Amount]", true)).toBe("'Book.xlsx'!Sales[Amount]");
	});

	it('escapes column punctuation and preserves special item selectors', () => {
		expect(
			renameTableColumnInFormula("Sales['#All]+Sales[#All]", 'Sales', '#All', "New [#@']"),
		).toBe("Sales[New '['#'@''']]+Sales[#All]");
		expect(renameTableColumnInFormula("Sales[A']B]", 'Sales', 'A]B', 'Amount')).toBe(
			'Sales[Amount]',
		);
	});
});
