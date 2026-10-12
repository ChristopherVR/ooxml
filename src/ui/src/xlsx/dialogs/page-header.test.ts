import { afterEach, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import {
	createTestContext,
	dialogEl,
	clickButton,
	inputByLabel,
	setValue,
} from '../commands/test-support';
import { registerToolDialogs } from './register-tools';

afterEach(() => {
	document.body.replaceChildren();
});

it('preserves native header codes on no-op OK and authors titles and fields with one undo', async () => {
	const wb = createWorkbook();
	const ctx = createTestContext(wb);
	const edit = ctx.session()!;
	const raw = '&L&"Calibri,Bold"Left &A&CCenter&R&P of &N';
	edit.setPageSetup(0, { header: raw });
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('page-setup');
	clickButton(dialogEl(ctx, 'page-setup'), 'OK');
	await done;
	expect(wb.sheets[0]!.pageSetup?.header).toBe(raw);
	const change = ctx.dialogs.open('page-setup');
	const dialog = dialogEl(ctx, 'page-setup');
	setValue(inputByLabel(dialog, 'Header:'), '&L&A&RPage &P of &N');
	setValue(inputByLabel(dialog, 'Rows to repeat at top'), '$1:$2');
	setValue(inputByLabel(dialog, 'Columns to repeat at left'), '$A:$A');
	clickButton(dialog, 'OK');
	await change;
	expect(wb.definedNames.some((name) => name.name === '_xlnm.Print_Titles')).toBe(true);
	edit.undo();
	expect(wb.sheets[0]!.pageSetup?.header).toBe(raw);
	expect(wb.definedNames).toEqual([]);
});
