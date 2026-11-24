// @vitest-environment jsdom
import { createWorkbook, DEFAULT_ERROR_RULES, type ErrorRule } from 'ooxml-core/xlsx';
import { auditState, setErrorRules } from 'ooxml-core/xlsx/ui';
import { afterEach, describe, expect, it } from 'vitest';
import { createTestContext } from './commands/test-support';
import {
	clearSavedErrorRules,
	ERROR_RULES_KEY,
	savedErrorRules,
	saveErrorRules,
} from './error-rules';
import { GridView } from './grid/grid-view';

afterEach(() => {
	clearSavedErrorRules();
	document.body.innerHTML = '';
});

describe('Excel Error Checking rule persistence', () => {
	it('defaults to Excel defaults and persists toggled rules in localStorage', () => {
		expect(savedErrorRules()).toEqual([...DEFAULT_ERROR_RULES]);

		const custom: ErrorRule[] = ['error-value', 'empty-cell-reference'];
		saveErrorRules(custom);
		expect(localStorage.getItem(ERROR_RULES_KEY)).toBe(JSON.stringify(custom));
		expect(savedErrorRules()).toEqual(custom);
	});

	it('ignores unknown rule names and malformed storage values', () => {
		localStorage.setItem(ERROR_RULES_KEY, JSON.stringify(['number-as-text', 'not-a-rule', 42]));
		expect(savedErrorRules()).toEqual(['number-as-text']);

		localStorage.setItem(ERROR_RULES_KEY, '{not-json');
		expect(savedErrorRules()).toEqual([...DEFAULT_ERROR_RULES]);
	});
});

describe('Excel background error-checking green triangles on the grid', () => {
	it('marks cells matching enabled rules, reacts to rule toggles, and respects ignoredErrors', () => {
		const wb = createWorkbook({ sheets: ['Sheet1'] });
		const ctx = createTestContext(wb);
		const session = ctx.session()!;
		// A1: number stored as text, B1: formula error, C1: normal number
		session.setCellInput(0, 0, 0, "'123");
		session.setCellInput(0, 0, 1, '=1/0');
		session.setCellInput(0, 0, 2, '42');
		const host = document.createElement('div');
		document.body.append(host);
		const view = new GridView(ctx, host);
		view.paint();

		const hasMark = (row: number, col: number): boolean =>
			Boolean(view.quads.main.cells.nodeAt(row, col)?.querySelector('.xg-er'));

		// Both A1 (number-as-text) and B1 (error-value) get green triangles; C1 does not.
		expect(hasMark(0, 0)).toBe(true);
		expect(hasMark(0, 1)).toBe(true);
		expect(hasMark(0, 2)).toBe(false);

		// Turning off number-as-text removes A1's triangle while keeping B1's.
		setErrorRules(
			ctx,
			auditState(ctx).rules.filter((r) => r !== 'number-as-text'),
		);
		view.paint();
		expect(hasMark(0, 0)).toBe(false);
		expect(hasMark(0, 1)).toBe(true);

		// Ignoring the error on B1 removes its triangle too.
		ctx.session()!.setIgnoredErrors(0, [
			{
				ranges: [{ start: { row: 0, col: 1 }, end: { row: 0, col: 1 } }],
				flags: ['evalError'],
			},
		]);
		view.invalidate();
		view.paint();
		expect(hasMark(0, 1)).toBe(false);

		view.destroy();
	});
});
