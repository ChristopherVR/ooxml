import { expect, test, type Locator, type Page } from '@playwright/test';
import type { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { savedDocument } from './saved-document';
import { fileInput } from './helpers';
import { controlEditsFixture } from './support/control-edits-fixture';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function open(page: Page) {
	await page.goto('/?framework=vanilla');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'control-edits.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: Buffer.from(await controlEditsFixture()),
	});
	const editor = page.locator('docx-editor');
	await expect(editor.locator('.ProseMirror td[data-sdt-cell]')).toHaveCount(1);
	return editor;
}

/**
 * Selects `text` in the editor (or puts the caret after it) and focuses it, then lets the editor's
 * check after focus run, so the keys pressed next act on that selection.
 */
async function select(editor: Locator, text: string, caret = false) {
	await editor.evaluate(
		(element, [text, caret]) => {
			const view = (element as unknown as { view: EditorView }).view;
			let from = -1;
			view.state.doc.descendants((node, pos) => {
				if (from < 0 && node.isText && node.text!.includes(text))
					from = pos + node.text!.indexOf(text);
			});
			const Selection = view.state.selection.constructor as typeof TextSelection;
			const to = from + text.length;
			view.dispatch(
				view.state.tr.setSelection(Selection.create(view.state.doc, caret ? to : from, to)),
			);
			view.focus();
		},
		[text, caret] as const,
	);
	await editor.page().evaluate(() => new Promise<void>((resolve) => setTimeout(resolve, 50)));
}

test('a pick in a cell control takes the locked control around its text away', async ({ page }) => {
	const editor = await open(page);
	const cell = editor.locator('.ProseMirror td[data-sdt-cell]');
	await expect(cell.locator('p[data-sdt-block="92"]')).toHaveText('Fixed');
	await cell.click();
	await editor.getByRole('option', { name: 'Secondary', exact: true }).click();
	await expect(cell).toHaveText('Secondary');
	await expect(cell.locator('p[data-sdt-block]')).toHaveCount(0);
	const { xml } = await savedDocument(page, 'control-edits-pick.docx');
	expect(xml).toContain('<w:t>Secondary</w:t>');
	expect(xml).not.toContain('<w:alias w:val="Fixed note"/>');
});

test('text typed after deleting all of a temporary control is not in a control', async ({
	page,
}) => {
	const editor = await open(page);
	const draft = editor.locator('.ProseMirror .dve-sdt[data-sdt-title="Draft"]');
	await expect(draft).toHaveText('Draft');
	await select(editor, 'Draft');
	await page.keyboard.press('Backspace');
	await page.keyboard.type('Z');
	await expect(editor.locator('.ProseMirror p', { hasText: 'Note:' })).toHaveText('Note: Z');
	await expect(editor.locator('.ProseMirror .dve-sdt')).toHaveCount(0);
	const { xml } = await savedDocument(page, 'control-edits-draft.docx');
	expect(xml).toMatch(/Z<\/w:t>/);
	expect(xml).not.toContain('<w:alias w:val="Draft"/>');
});

test('typing in a table inside a temporary control removes all of the control', async ({
	page,
}) => {
	const editor = await open(page);
	const section = editor.locator('.ProseMirror p[data-sdt-block="94"]');
	await expect(section).toHaveCount(2);
	await select(editor, 'Inside', true);
	await page.keyboard.type('!');
	await expect(editor.locator('.ProseMirror td', { hasText: 'Inside' })).toHaveText('Inside!');
	await expect(section).toHaveCount(0);
	const { xml } = await savedDocument(page, 'control-edits-section.docx');
	expect(xml).toContain('<w:t>Inside!</w:t>');
	expect(xml).not.toContain('<w:alias w:val="Section"/>');
});

/** Pastes a table from another page at the end of the document: one cell, with `control` on it. */
async function pasteCellFromPage(editor: Locator, control: object, text: string) {
	await editor.evaluate(
		(element, [control, text]) => {
			const view = (element as unknown as { view: EditorView }).view;
			const Selection = view.state.selection.constructor as typeof TextSelection;
			view.dispatch(
				view.state.tr.setSelection(
					Selection.create(view.state.doc, view.state.doc.content.size - 1),
				),
			);
			view.focus();
			const [table, body, row, cell, paragraph] = ['table', 'tbody', 'tr', 'td', 'p'].map((tag) =>
				document.createElement(tag),
			);
			cell!.dataset.sdt = JSON.stringify(control);
			paragraph!.textContent = text;
			cell!.append(paragraph!);
			row!.append(cell!);
			body!.append(row!);
			table!.append(body!);
			// Marked as a whole table, as the editor marks its own copies; a page can do the same.
			table!.dataset.pmSlice = '0 0 []';
			const data = new DataTransfer();
			data.setData('text/html', table!.outerHTML);
			view.dom.dispatchEvent(
				new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
			);
		},
		[control, text] as const,
	);
}

test('a cell pasted from another page brings no content in its control properties', async ({
	page,
}) => {
	const editor = await open(page);
	// The page put a field that fetches a picture where the control's properties go.
	await pasteCellFromPage(
		editor,
		{
			id: '90',
			level: 'cell',
			kind: 'richText',
			sourceXml: `<w:p xmlns:w="${W}"><w:fldSimple w:instr='INCLUDEPICTURE "http://example.test/a.png"'/></w:p>`,
		},
		'Pasted',
	);
	await expect(editor.locator('.ProseMirror td[data-sdt-cell="90"]')).toHaveText('Pasted');
	const { xml } = await savedDocument(page, 'control-edits-paste.docx');
	expect(xml).toContain('<w:t>Pasted</w:t>');
	expect(xml).toMatch(/<w:sdt><w:sdtPr><w:id w:val="90"\/><\/w:sdtPr><w:sdtContent><w:tc>/);
	expect(xml).not.toContain('INCLUDEPICTURE');
});

test('a cell pasted from another page keeps none of the properties it had there', async ({
	page,
}) => {
	const editor = await open(page);
	// Properties Word reads, but another document's: they bind the cell to that page's data.
	await pasteCellFromPage(
		editor,
		{
			id: '96',
			level: 'cell',
			kind: 'richText',
			alias: 'Theirs',
			sourceXml:
				`<w:sdtPr xmlns:w="${W}"><w:alias w:val="Theirs"/><w:id w:val="96"/>` +
				`<w:dataBinding w:xpath="/theirs" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/></w:sdtPr>`,
			endSourceXml: `<w:sdtEndPr xmlns:w="${W}"><w:rPr><w:b/></w:rPr></w:sdtEndPr>`,
		},
		'Theirs',
	);
	await expect(editor.locator('.ProseMirror td[data-sdt-cell="96"]')).toHaveText('Theirs');
	const { xml } = await savedDocument(page, 'control-edits-paste-bound.docx');
	expect(xml).toMatch(
		/<w:sdt><w:sdtPr><w:alias w:val="Theirs"\/><w:id w:val="96"\/><\/w:sdtPr><w:sdtContent><w:tc>/,
	);
	expect(xml).not.toContain('/theirs');
	expect(xml).not.toContain('w:sdtEndPr');
});
