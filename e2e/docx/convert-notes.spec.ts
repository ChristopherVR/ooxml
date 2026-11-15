import { expect, test } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument, dialogByHeading } from './helpers';

const note = (prefix: string, kind: 'footnote' | 'endnote', text: string) => ({
	id: '1',
	blocks: [
		{
			type: 'paragraph',
			id: `${prefix}1-p0`,
			runs: [{ text: '', noteMark: kind, verticalAlign: 'superscript' }, { text }],
		},
	],
});

test('Convert Notes swaps footnotes and endnotes, and a right-click converts one note', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 900 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate(
		(el, notes) => {
			const target = el as DocxEditorElement;
			const model = structuredClone(target.documentModel!);
			model.blocks = [
				{
					type: 'paragraph',
					id: 'p1',
					runs: [
						{ text: 'Alpha' },
						{
							text: '',
							noteReference: { kind: 'footnote', id: '1' },
							verticalAlign: 'superscript',
						},
						{ text: ' Beta' },
						{ text: '', noteReference: { kind: 'endnote', id: '1' }, verticalAlign: 'superscript' },
						{ text: ' end' },
					],
				},
			];
			model.footnotes = [notes.foot] as never;
			model.endnotes = [notes.end] as never;
			target.documentModel = model;
		},
		{ foot: note('fn', 'footnote', ' Foot text'), end: note('en', 'endnote', ' End text') },
	);
	const body = editor.locator('.dve-paper > .ProseMirror');
	const references = body.locator('[data-docx-note-kind]');
	await expect(references).toHaveCount(2);
	await editor.getByRole('tab', { name: 'References', exact: true }).click();
	await editor.getByRole('button', { name: 'Convert notes', exact: true }).click();
	const dialog = dialogByHeading(editor, 'Convert Notes');
	await expect(dialog).toBeVisible();
	await dialog.getByRole('radio', { name: 'Swap footnotes and endnotes', exact: true }).check();
	await dialog.getByRole('button', { name: 'OK', exact: true }).click();
	await expect(dialog).toBeHidden();
	await expect(references.nth(0)).toHaveAttribute('data-docx-note-kind', 'endnote');
	await expect(references.nth(1)).toHaveAttribute('data-docx-note-kind', 'footnote');
	await expect(editor.locator('.dve-notes-footnote')).toContainText('End text');
	await expect(editor.locator('.dve-notes-endnote')).toContainText('Foot text');

	await references.nth(1).click({ button: 'right' });
	const menu = editor.getByRole('menu');
	await expect(menu).toBeVisible();
	await menu.getByRole('menuitem', { name: 'Convert to Endnote', exact: true }).click();
	await expect(references.nth(1)).toHaveAttribute('data-docx-note-kind', 'endnote');
	const model = await editor.evaluate((el) => (el as DocxEditorElement).documentModel!);
	expect(model.footnotes).toEqual([]);
	expect(model.endnotes).toHaveLength(2);

	await body.press('Control+z');
	await expect(references.nth(1)).toHaveAttribute('data-docx-note-kind', 'footnote');
});
