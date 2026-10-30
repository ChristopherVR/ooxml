// @vitest-environment jsdom
import { createWorkbook } from 'ooxml-core/xlsx';
import { auditState, formulaCommands } from 'ooxml-core/xlsx/ui';
import { afterEach, describe, expect, it } from 'vitest';
import {
	clickButton,
	createTestContext,
	dialogEl,
	type TestContext,
} from '../commands/test-support';
import { registerNavigationDialogs } from './register-navigation';
import { mountWatchWindow } from './watch-window';

afterEach(() => (document.body.innerHTML = ''));

function setup(): TestContext {
	const ctx = createTestContext(createWorkbook({ sheets: ['Sheet1'] }));
	registerNavigationDialogs(ctx);
	ctx.commands.registerAll(formulaCommands());
	const s = ctx.session()!;
	s.setCellInput(0, 0, 0, '2');
	s.setCellInput(0, 1, 0, '3');
	s.setCellInput(0, 0, 1, '=A1+A2');
	s.setCellInput(0, 0, 2, '=B1*2');
	return ctx;
}

describe('formula auditing', () => {
	it('traces precedents level by level and removes them', async () => {
		const ctx = setup();
		ctx.select('C1');
		await ctx.commands.run('formulas.trace-precedents');
		expect(auditState(ctx).arrows).toHaveLength(1);
		await ctx.commands.run('formulas.trace-precedents');
		expect(auditState(ctx).arrows).toHaveLength(3);
		await ctx.commands.run('formulas.remove-precedent-arrows');
		expect(auditState(ctx).arrows).toHaveLength(1);
		await ctx.commands.run('formulas.remove-arrows');
		expect(auditState(ctx).arrows).toHaveLength(0);
	});

	it('tells the user when the cell has nothing to trace', async () => {
		const ctx = setup();
		ctx.select('A1');
		await ctx.commands.run('formulas.trace-precedents');
		expect(ctx.toasts.at(-1)?.message).toContain('Trace Precedents');
	});

	it('steps through Evaluate Formula', async () => {
		const ctx = setup();
		ctx.select('B1');
		const done = ctx.dialogs.open('evaluate-formula');
		const dialog = dialogEl(ctx, 'evaluate-formula');
		expect(dialog.querySelector('u')?.textContent).toBe('A1');
		clickButton(dialog, 'Evaluate');
		expect(dialog.textContent).toContain('=2+A2');
		clickButton(dialog, 'Evaluate');
		clickButton(dialog, 'Evaluate');
		expect(dialog.textContent).toContain('=5');
		clickButton(dialog, 'Restart');
		expect(dialog.textContent).toContain('=A1+A2');
		clickButton(dialog, 'Close');
		await done;
	});

	it('reports an Evaluate Formula request on a plain cell', async () => {
		const ctx = setup();
		ctx.select('A1');
		await ctx.dialogs.open('evaluate-formula');
		expect(ctx.toasts.at(-1)?.message).toContain('Evaluate Formula');
	});

	it('walks the error checking issues', async () => {
		const ctx = setup();
		ctx.session()!.setCellInput(0, 5, 0, '=1/0');
		ctx.select('A1');
		const done = ctx.dialogs.open('error-checking');
		const dialog = dialogEl(ctx, 'error-checking');
		expect(dialog.textContent).toContain('Error in cell A6');
		clickButton(dialog, 'Ignore Error');
		await done;
		expect(ctx.toasts.at(-1)?.message).toContain('error check is complete');
	});

	it('lists watched cells and drops them', async () => {
		const ctx = setup();
		mountWatchWindow(ctx);
		await ctx.commands.run('formulas.watch-window');
		const panel = ctx.root.querySelector<HTMLElement>('[data-part="watch-window"]')!;
		expect(panel.hidden).toBe(false);
		const { addWatch } = await import('ooxml-core/xlsx/ui');
		addWatch(ctx, { sheet: 0, row: 0, col: 1 });
		expect(panel.textContent).toContain('=A1+A2');
		expect(panel.textContent).toContain('5');
		panel.querySelector<HTMLElement>('tbody tr')!.click();
		clickButton(panel, 'Delete Watch');
		expect(auditState(ctx).watches).toHaveLength(0);
		await ctx.commands.run('formulas.watch-window');
		expect(panel.hidden).toBe(true);
	});
});
