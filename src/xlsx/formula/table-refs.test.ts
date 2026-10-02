import { describe, expect, it } from 'vitest';
import { renameTableInFormula } from './table-refs.js';

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
