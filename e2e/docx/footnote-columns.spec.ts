import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { clickInEditor, fileInput, reveal, dialogByHeading } from './helpers';
test('native footnote columns render, edit as one history step and save/reopen', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	const source = readFileSync(
		new URL(
			'../../src/core/docx/__fixtures__/footnote-columns/two-two-twenty.docx',
			import.meta.url,
		),
	);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'native-columns.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: source,
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Left body');
	const print = async () => {
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	};
	await print();
	const notes = editor.locator('.dve-print-footnote');
	await expect(notes).toHaveCount(2);
	await expect(notes.nth(1)).toHaveCSS('left', '237.333px');
	await expect(notes.first()).toContainText('N0L09');
	await expect(notes.nth(1)).toContainText('N0L10');
	await editor.getByRole('tab', { name: 'Layout', exact: true }).click();
	const columns = editor.getByRole('combobox', { name: 'Columns', exact: true });
	await reveal(editor, columns);
	await columns.press('Alt+ArrowDown');
	await editor.getByRole('menuitem', { name: 'More Columns…', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Columns');
	await dialog
		.getByRole('spinbutton', { name: 'Footnote columns (0 matches section)', exact: true })
		.fill('3');
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(notes).toHaveCount(3);
	await page.keyboard.press('Control+z');
	await expect(notes).toHaveCount(2);
	await page.keyboard.press('Control+y');
	await expect(notes).toHaveCount(3);
	const bytes = await editor.evaluate(async (node) =>
		Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
	);
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'reopened.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: Buffer.from(bytes),
	});
	await print();
	await expect(notes).toHaveCount(3);
	await expect(notes.nth(2)).toContainText('N0L19');
});

for (const mode of [12, 14])
	test(`legacy mode ${mode} retains reference attachment and per-column continuation separators after edit/undo/save`, async ({
		page,
	}) => {
		await page.goto('/?framework=vanilla');
		const source = readFileSync(
			new URL(
				'../../src/core/docx/__fixtures__/footnote-columns/mode' + mode + '-two-custom-eighty.docx',
				import.meta.url,
			),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'legacy-notes.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: source,
		});
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror').first()).toContainText('Left body');
		const layout = editor.getByRole('combobox', { name: 'Layout view', exact: true });
		const print = async () => {
			await editor.getByRole('tab', { name: 'View', exact: true }).click();
			await layout.selectOption('print');
		};
		const verify = async () => {
			const pages = editor.locator('.dve-print-page');
			await expect(pages).toHaveCount(2);
			await expect(pages.first().locator('.dve-print-footnote')).toHaveCount(1);
			await expect(pages.first().locator('.dve-print-footnote')).toHaveCSS('left', '226.667px');
			await expect(pages.first().locator('.dve-print-footnote .dve-print-line')).toHaveCount(33);
			const next = pages.nth(1).locator('.dve-print-footnote');
			await expect(next).toHaveCount(2);
			await expect(next.first().locator('.dve-print-line')).toHaveCount(34);
			await expect(next.nth(1).locator('.dve-print-line')).toHaveCount(13);
			await expect(next.nth(1)).toHaveCSS('top', '336px');
			await expect(pages.first().locator('.dve-print-custom-footnote-separator')).toContainText(
				'Sources',
			);
			const separators = pages.nth(1).locator('.dve-print-custom-footnote-separator');
			await expect(separators).toHaveCount(2);
			await expect(separators.first()).toContainText('Continued');
			await expect(separators.nth(1)).toContainText('Continued');
			await expect(separators.nth(1)).toHaveCSS('top', '320px');
		};
		await print();
		await verify();
		await layout.selectOption('draft');
		const body = editor.locator('.ProseMirror').first();
		await clickInEditor(body);
		await page.keyboard.press('Control+Home');
		await page.keyboard.type('Prefix ');
		await expect(body).toContainText('Prefix Left');
		await page.keyboard.press('Control+z');
		await expect(body).not.toContainText('Prefix');
		await page.keyboard.press('Control+y');
		await expect(body).toContainText('Prefix Left');
		await print();
		await verify();
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'legacy-reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await print();
		await verify();
	});

test('legacy partial note continuation leaves following body text in the second column', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'legacy-partial.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: readFileSync(
			new URL(
				'../../src/core/docx/__fixtures__/footnote-columns/mode14-left-custom-sixty.docx',
				import.meta.url,
			),
		),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror').first()).toContainText('Right body');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
	await expect(editor.locator('.dve-print-page')).toHaveCount(1);
	const columns = editor.locator('.dve-print-column');
	await expect(columns.nth(1)).toContainText('Right body text');
	const notes = editor.locator('.dve-print-footnote');
	await expect(notes).toHaveCount(2);
	await expect(notes.first().locator('.dve-print-line')).toHaveCount(33);
	await expect(notes.nth(1).locator('.dve-print-line')).toHaveCount(27);
	const separators = editor.locator('.dve-print-custom-footnote-separator');
	await expect(separators).toHaveCount(2);
	await expect(separators.first()).toContainText('Sources');
	await expect(separators.nth(1)).toContainText('Continued');
});
