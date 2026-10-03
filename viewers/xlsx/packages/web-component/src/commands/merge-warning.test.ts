// @vitest-environment jsdom
// Merging over several values asks first (Excel's warning, OK / Cancel); a merge that loses
// nothing does not ask.
import { createWorkbook, getCell } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { registerSimpleDialogs } from '../dialogs/simple.js';
import { allCommands } from './index.js';
import { clickButton, createTestContext, dialogEl, tick } from './test-support.js';

afterEach(() => (document.body.innerHTML = ''));

function setup(values: string[]) {
	const ctx = createTestContext(createWorkbook());
	ctx.commands.registerAll(allCommands());
	registerSimpleDialogs(ctx);
	values.forEach((value, col) => ctx.session()!.setCellInput(0, 0, col, value));
	ctx.select('A1:C1');
	return ctx;
}
const merges = (ctx: ReturnType<typeof setup>) => ctx.workbook()!.sheets[0]!.merges;
const WARNING = 'Merging cells only keeps the upper-left value and discards other values.';

describe('merge warning', () => {
	it('asks before discarding values and merges on OK', async () => {
		const ctx = setup(['a', 'b']);
		const run = ctx.commands.run('home.merge-center');
		await tick();
		const dialog = dialogEl(ctx, 'confirm');
		expect(dialog.textContent).toContain(WARNING);
		clickButton(dialog, 'OK');
		await run;
		expect(merges(ctx)).toHaveLength(1);
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 1)?.value ?? null).toBeNull();
	});

	it('keeps the cells unmerged on Cancel', async () => {
		const ctx = setup(['a', 'b']);
		const run = ctx.commands.run('home.merge-cells');
		await tick();
		clickButton(dialogEl(ctx, 'confirm'), 'Cancel');
		await run;
		expect(merges(ctx)).toHaveLength(0);
		expect(getCell(ctx.workbook()!.sheets[0]!, 0, 1)?.value).toBe('b');
	});

	it('merges a single value without asking', async () => {
		const ctx = setup(['only']);
		await ctx.commands.run('home.merge-center');
		expect(ctx.root.querySelector('[data-dialog="confirm"]')).toBeNull();
		expect(merges(ctx)).toHaveLength(1);
	});
});
