import { afterEach, expect, it } from 'vitest';
import { buildPrintHtml } from '../print-pages';
import { installCommands } from './index';
import { clickButton, createTestContext, dialogEl, inputByLabel } from './test-support';

afterEach(() => document.body.replaceChildren());

it('applies and undoes independent fit controls, then switches back to percentage scaling', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const session = ctx.session()!;
		await ctx.commands.run('page.fit-width', '3');
		await ctx.commands.run('page.fit-height', '2');
		await ctx.commands.run('page.scale', '80%');
		expect(session.workbook.sheets[0]?.pageSetup).toEqual({ scale: 80 });
		expect(session.undo()).toBe(true);
		expect(session.workbook.sheets[0]?.pageSetup).toEqual({ fitToWidth: 3, fitToHeight: 2 });
		await ctx.commands.run('page.fit-height', '0');
		expect(session.workbook.sheets[0]?.pageSetup).toEqual({ fitToWidth: 3, fitToHeight: 0 });
	} finally {
		stop();
	}
});

it('opens Print Titles on the Sheet tab and edits the scoped name in one undo step', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		await ctx.commands.run('page.print-titles');
		const dialog = dialogEl(ctx, 'page-setup');
		expect(dialog.querySelector('[aria-selected="true"]')?.textContent).toBe('Sheet');
		const rows = inputByLabel<HTMLInputElement>(dialog, 'Rows to repeat at top');
		rows.value = '$1:$2';
		rows.dispatchEvent(new Event('input', { bubbles: true }));
		clickButton(dialog, 'OK');
		expect(ctx.workbook()?.definedNames[0]?.name).toBe('_xlnm.Print_Titles');
		expect(ctx.session()?.undo()).toBe(true);
		expect(ctx.workbook()?.definedNames).toEqual([]);
	} finally {
		stop();
	}
});

it('creates a legacy note and switches to Page Layout with undo', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const open = ctx.commands.run('insert.note');
		const dialog = dialogEl(ctx, 'note');
		inputByLabel<HTMLTextAreaElement>(dialog, 'Note').value = 'Review the assumptions';
		clickButton(dialog, 'OK');
		expect(await open).toBe(true);
		const sheet = ctx.workbook()?.sheets[0];
		expect(sheet?.comments[0]?.text).toBe('Review the assumptions');
		expect(sheet?.comments[0]?.replies).toBeUndefined();
		await ctx.commands.run('insert.header-footer');
		expect(sheet?.view.mode).toBe('pageLayout');
		expect(ctx.session()?.undo()).toBe(true);
		expect(sheet?.view.mode).toBeUndefined();
		expect(ctx.session()?.undo()).toBe(true);
		expect(sheet?.comments).toEqual([]);
	} finally {
		stop();
	}
});

it('deletes a background with undo and keeps it out of the print document', async () => {
	const ctx = createTestContext();
	const stop = installCommands(ctx);
	try {
		const session = ctx.session()!;
		session.setCellInput(0, 0, 0, 'Report');
		session.setSheetBackground(0, {
			bytes: Uint8Array.from([137, 80, 78, 71]),
			contentType: 'image/png',
		});
		expect(buildPrintHtml(session.workbook, 0, 'Report')).not.toContain('data:image/');
		expect(await ctx.commands.run('page.background-delete')).toBe(true);
		expect(session.workbook.sheets[0]?.backgroundPicture).toBeUndefined();
		expect(session.undo()).toBe(true);
		expect(session.workbook.sheets[0]?.backgroundPicture).toBeDefined();
	} finally {
		stop();
	}
});
