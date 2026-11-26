import { expect, it, vi } from 'vitest';
import { pendingXPath } from '#xpath-runtime';
import { parseRange } from '../../address';
import { createEditSession } from '../../edit/session';
import { createConditionalFormatEvaluator } from '../../layout/cf-evaluator';
import { loadXlsx } from '../../read/load';
import { saveXlsx } from '../../write/save';
import { book, engine, get } from '../test-helpers';

const library = vi.hoisted(() => {
	let resolve: () => void = () => {};
	const promise = new Promise<void>((ready) => {
		resolve = ready;
	});
	return { promise, resolve };
});

vi.mock('xpath', async (importOriginal) => {
	await library.promise;
	return importOriginal();
});

it('prepares the first FILTERXML rule and refreshes formatting in manual calculation mode', async () => {
	const node = vi.spyOn(process, 'getBuiltinModule').mockReturnValue(undefined);
	try {
		const source = book({ A1: '<r><a>7</a></r>', B1: 1, C1: '=B1+1' });
		engine(source).recalculateAll();
		source.calcMode = 'manual';
		const workbook = await loadXlsx(await saveXlsx(source));
		const session = createEditSession(workbook);
		session.setCellInput(0, 0, 1, '9');
		const range = parseRange('B1')!;
		const style = { font: { bold: true } };
		const formatting = () =>
			createConditionalFormatEvaluator(workbook, 0, (formula, at) =>
				session.calc.evaluate(formula, at),
			).at(0, 1)?.style;
		const refreshed = vi.fn(formatting);
		session.onChange((change) => {
			if (change.viewOnly) refreshed();
		});

		expect(pendingXPath(workbook)).toBeUndefined();
		session.addConditionalFormat(0, {
			ranges: [range],
			rules: [{ type: 'expression', formula: 'FILTERXML($A$1,"//a")=7', priority: 1, style }],
		});
		const ready = pendingXPath(workbook);
		expect(ready).toBeDefined();
		expect(formatting()).toBeUndefined();

		// A host can also evaluate a rule that did not come through the rule editor.
		const other = createEditSession(book({ A1: '<r><a>7</a></r>' }));
		const adHoc = () => other.calc.evaluate('FILTERXML(A1,"//a")', { sheet: 0, row: 0, col: 1 });
		const adHocRefresh = vi.fn(adHoc);
		other.onChange((change) => {
			if (change.viewOnly) adHocRefresh();
		});
		expect(adHoc()).toEqual({ error: '#BUSY!' });

		library.resolve();
		await ready;
		expect(refreshed).toHaveReturnedTimes(1);
		expect(refreshed).toHaveReturnedWith(style);
		expect(adHocRefresh).toHaveReturnedWith(7);
		expect(get(workbook, 'C1')).toBe(2);
		expect(workbook.valuesStale).toBe(true);
		session.undo();
		expect(formatting()).toBeUndefined();
		session.redo();
		expect(formatting()).toEqual(style);
	} finally {
		library.resolve();
		node.mockRestore();
	}
});
