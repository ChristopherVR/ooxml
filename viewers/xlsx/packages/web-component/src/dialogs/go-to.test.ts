// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	inputByLabel,
	pressKey,
	setValue,
	tick,
} from '../commands/test-support.js';
import { resolveReference } from './go-to.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1', 'My Data'] }));
	registerNavigationDialogs(ctx);
	ctx.session()!.setDefinedName({ name: 'Totals', formula: "'My Data'!$B$2:$C$4" });
	return ctx;
}

describe('go to', () => {
	it('resolves references and names', () => {
		const ctx = setup();
		expect(resolveReference(ctx, 'B3')).toEqual({
			sheet: 0,
			range: { start: { row: 2, col: 1 }, end: { row: 2, col: 1 } },
		});
		expect(resolveReference(ctx, "'My Data'!$A$1:B2")?.sheet).toBe(1);
		expect(resolveReference(ctx, 'totals')).toEqual({
			sheet: 1,
			range: { start: { row: 1, col: 1 }, end: { row: 3, col: 2 } },
		});
		expect(resolveReference(ctx, 'Nope!A1')).toBeUndefined();
		expect(resolveReference(ctx, 'unknown')).toBeUndefined();
	});

	it('selects a typed reference', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('go-to');
		const dialog = dialogEl(ctx, 'go-to');
		setValue(inputByLabel(dialog, 'Reference:'), 'C5:D6');
		clickButton(dialog, 'OK');
		expect(await done).toBe('C5:D6');
		expect(ctx.selection.get().ranges[0]).toEqual({
			start: { row: 4, col: 2 },
			end: { row: 5, col: 3 },
		});
	});

	it('picks a defined name from the list and switches sheets', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('go-to');
		const dialog = dialogEl(ctx, 'go-to');
		(dialog.querySelector('[role="option"]') as HTMLElement).click();
		expect(inputByLabel(dialog, 'Reference:').value).toBe('Totals');
		clickButton(dialog, 'OK');
		await done;
		expect(ctx.activeSheet()).toBe(1);
		expect(ctx.selection.get().active).toEqual({ row: 1, col: 1 });
	});

	it('keeps the dialog open on a bad reference; Cancel changes nothing', async () => {
		const ctx = setup();
		const done = ctx.dialogs.open('go-to');
		const dialog = dialogEl(ctx, 'go-to');
		setValue(inputByLabel(dialog, 'Reference:'), '???');
		clickButton(dialog, 'OK');
		await tick();
		expect(ctx.toasts[0]?.message).toBe('Reference is not valid.');
		clickButton(dialog, 'Cancel');
		expect(await done).toBeUndefined();
		expect(ctx.selection.get().active).toEqual({ row: 0, col: 0 });
	});

	it('Special... opens Go To Special', async () => {
		const ctx = setup();
		void ctx.dialogs.open('go-to');
		clickButton(dialogEl(ctx, 'go-to'), 'Special...');
		await tick();
		expect(ctx.root.querySelector('[data-dialog="go-to"]')).toBeNull();
		expect(dialogEl(ctx, 'go-to-special')).toBeTruthy();
	});
});

describe('go to special', () => {
	it('selects formulas only of the checked types', async () => {
		const ctx = setup();
		const s = ctx.session()!;
		s.setCellInput(0, 0, 0, '1');
		s.setCellInput(0, 1, 0, '=A1+1');
		s.setCellInput(0, 2, 0, '="x"');
		const done = ctx.dialogs.open('go-to-special');
		const dialog = dialogEl(ctx, 'go-to-special');
		inputByLabel(dialog, 'Formulas').click();
		inputByLabel(dialog, 'Text').click();
		clickButton(dialog, 'OK');
		expect(await done).toBe('formulas');
		expect(ctx.selection.get().ranges).toEqual([
			{ start: { row: 1, col: 0 }, end: { row: 1, col: 0 } },
		]);
	});

	it('stays open with a toast when nothing matches; Escape closes', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 0, 0, '1');
		const done = ctx.dialogs.open('go-to-special');
		const dialog = dialogEl(ctx, 'go-to-special');
		inputByLabel(dialog, 'Notes').click();
		clickButton(dialog, 'OK');
		await tick();
		expect(ctx.toasts[0]?.message).toBe('No cells were found.');
		pressKey(inputByLabel(dialog, 'Notes'), 'Escape');
		expect(await done).toBeUndefined();
	});
});
