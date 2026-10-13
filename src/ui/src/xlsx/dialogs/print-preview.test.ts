import { afterEach, expect, it } from 'vitest';
import { createWorkbook } from 'ooxml-core/xlsx';
import { createTestContext, dialogEl, clickButton } from '../commands/test-support';
import { registerToolDialogs } from './register-tools';

afterEach(() => {
	document.body.replaceChildren();
});

it('previews core pages with navigation and reports unsupported print features', async () => {
	const ctx = createTestContext(createWorkbook());
	const edit = ctx.session()!;
	edit.setCellInput(0, 0, 0, 'First');
	edit.setCellInput(0, 8, 0, 'Last');
	edit.setPageBreak(0, 'row', 5);
	edit.setPageSetup(0, { header: '&G' });
	registerToolDialogs(ctx);
	const done = ctx.dialogs.open('print-preview');
	const dialog = dialogEl(ctx, 'print-preview');
	expect(dialog.textContent).toContain('Pages: 2 | Scale: 100%');
	expect(dialog.textContent).toContain(
		'Unsupported header/footer codes are preserved but not rendered.',
	);
	clickButton(dialog, 'Next Page');
	expect(dialog.textContent).toContain('Page 2 of 2');
	clickButton(dialog, 'Previous Page');
	expect(dialog.textContent).toContain('Page 1 of 2');
	clickButton(dialog, 'Actual Size');
	expect(dialog.textContent).toContain('Fit Page');
	clickButton(dialog, 'Close');
	await done;
});
