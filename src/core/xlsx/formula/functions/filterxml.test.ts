import { describe, expect, it } from 'vitest';
import corpus from '../__fixtures__/filterxml-native.json';
import { createEditSession } from '../../edit/session';
import { loadXlsx } from '../../read/index';
import { saveXlsx } from '../../write/index';
import { book, engine, get, grid } from '../test-helpers';

describe('FILTERXML native Excel results', () => {
	for (const [index, item] of corpus.entries()) {
		it(`${index + 1}: ${item.xpath}`, () => {
			const wb = book({ A1: item.xml, B1: item.xpath, C1: item.formula });
			engine(wb).recalculateAll();
			expect(grid(wb, `C1:C${item.values.length}`)).toEqual(item.values.map((value) => [value]));
		});
	}
	it('recalculates changed XML, undoes, and saves spilled results', async () => {
		const wb = book({ A1: '<r><a>1</a><a>2</a></r>', B1: '//a', C1: '=FILTERXML(A1,B1)' });
		engine(wb).recalculateAll();
		const edit = createEditSession(wb);
		expect(grid(wb, 'C1:C2')).toEqual([[1], [2]]);
		edit.setCellValue(0, 0, 0, '<r><a>changed</a></r>');
		expect(grid(wb, 'C1:C2')).toEqual([['changed'], [null]]);
		edit.undo();
		expect(grid(wb, 'C1:C2')).toEqual([[1], [2]]);
		const again = await loadXlsx(await saveXlsx(wb));
		engine(again).recalculateAll();
		expect(grid(again, 'C1:C2')).toEqual([[1], [2]]);
	});
	it('propagates input errors and respects spill obstructions', () => {
		const wb = book({
			A1: '<r><a>1</a><a>2</a></r>',
			C1: '=FILTERXML(A1,"//a")',
			C2: 'blocked',
			D1: '=FILTERXML(1/0,"//a")',
		});
		engine(wb).recalculateAll();
		expect(get(wb, 'C1')).toEqual({ error: '#SPILL!' });
		expect(get(wb, 'D1')).toEqual({ error: '#DIV/0!' });
	});
});
