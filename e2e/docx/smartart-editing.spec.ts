import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { savedDocument } from './saved-document';
import { dialogByHeading, fileInput, newDocument, reopenDocument, reveal } from './helpers';

const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
// Word's own output (Basic Block List, Alpha..Epsilon); provenance in src/core PROVENANCE.md.
const WORD_FIXTURE = new URL('./support/smartart-word.docx', import.meta.url);

const pane = (editor: Locator) => editor.locator('office-ui-smartart-text-pane');
const lines = (editor: Locator) => pane(editor).locator('input[data-node]');
const texts = (editor: Locator) =>
	lines(editor).evaluateAll((inputs) => inputs.map((input) => (input as HTMLInputElement).value));
const diagrams = (editor: Locator) => editor.locator('.ProseMirror .dve-smartart');

async function open(page: Page, bytes: Uint8Array | Buffer, name: string) {
	await (await fileInput(page)).setInputFiles({ name, mimeType: DOCX, buffer: Buffer.from(bytes) });
}

async function designButton(editor: Locator, name: string) {
	const button = editor.getByRole('button', { name, exact: true });
	await reveal(editor, button);
	return button;
}

/** Opens the text pane from the SmartArt Design tab (the diagram must be selected). */
async function openTextPane(editor: Locator) {
	await editor.getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
	await (await designButton(editor, 'Text Pane')).click();
	await expect(pane(editor)).toBeVisible();
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Insert > SmartArt, edit nodes in the text pane, save and reload`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1700, height: 950 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.locator('.ProseMirror').click();
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const insert = editor.getByRole('button', { name: 'Insert SmartArt', exact: true });
		await reveal(editor, insert);
		await insert.click();

		const dialog = dialogByHeading(editor, 'Choose a SmartArt Graphic');
		const picker = dialog.locator('office-ui-smartart-picker');
		await expect(picker.getByRole('option', { name: 'Basic Block List' })).toBeVisible();
		await picker.getByRole('button', { name: 'Process', exact: true }).click();
		await picker.getByRole('button', { name: 'List', exact: true }).click();
		await picker.getByRole('option', { name: 'Basic Block List' }).click();
		await dialog.getByRole('button', { name: 'OK', exact: true }).click();

		// The new diagram is selected, drawn approximately, and the contextual tab appears.
		await expect(diagrams(editor)).toHaveCount(1);
		const diagram = diagrams(editor).first();
		await expect(diagram).toHaveClass(/dve-smartart-selected/);
		await expect(diagram).toHaveAttribute('data-rendering', 'interpreted');
		await expect(editor.getByRole('tab', { name: 'SmartArt Design', exact: true })).toBeVisible();
		const shapes = () =>
			diagram.locator('office-ui-smartart').evaluate((element) => {
				return element.shadowRoot!.querySelectorAll('svg > g').length;
			});
		await expect.poll(shapes).toBe(3);

		// Text pane: type, Enter adds after, Tab demotes, Shift+Tab promotes, Backspace removes.
		await openTextPane(editor);
		await expect(lines(editor)).toHaveCount(3);
		await lines(editor).nth(0).fill('Plan');
		await lines(editor).nth(0).press('Enter');
		await expect(lines(editor)).toHaveCount(4);
		await expect(lines(editor).nth(1)).toBeFocused();
		await page.keyboard.type('Design');
		await page.keyboard.press('Tab');
		await expect(pane(editor).locator('[role="treeitem"]').nth(1)).toHaveAttribute(
			'aria-level',
			'2',
		);
		await page.keyboard.press('Shift+Tab');
		await expect(pane(editor).locator('[role="treeitem"]').nth(1)).toHaveAttribute(
			'aria-level',
			'1',
		);
		await lines(editor).nth(2).fill('Build');
		await lines(editor).nth(3).click();
		await page.keyboard.press('Backspace');
		await expect(lines(editor)).toHaveCount(3);
		await expect.poll(() => texts(editor)).toEqual(['Plan', 'Design', 'Build']);
		await expect(diagram.locator('office-ui-smartart')).toBeAttached();
		const drawn = () =>
			diagrams(editor)
				.first()
				.locator('office-ui-smartart')
				.evaluate((element) =>
					[...element.shadowRoot!.querySelectorAll('svg text')].map((text) => text.textContent),
				);
		await expect.poll(drawn).toEqual(['Plan', 'Design', 'Build']);

		// Ribbon: Demote acts on the text pane's current node (Build), which goes under Design.
		await editor.getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
		await (await designButton(editor, 'Demote')).click();
		await expect.poll(() => texts(editor)).toEqual(['Plan', 'Design', 'Build']);
		await editor.getByRole('button', { name: 'SmartArt Styles', exact: true }).click();
		await editor.getByRole('button', { name: 'Moderate Effect', exact: true }).click();
		await editor.getByRole('button', { name: 'Change Colors', exact: true }).click();
		await editor.getByRole('button', { name: 'Colored Fill - Accent 2', exact: true }).click();

		const { bytes } = await savedDocument(page, `ui-smartart-${framework}.docx`);
		const zip = await JSZip.loadAsync(bytes);
		const data = Object.keys(zip.files).find((name) => /word\/diagrams\/data.*\.xml$/.test(name));
		expect(data).toBeTruthy();
		expect(await zip.file(data!)!.async('string')).toContain(
			'urn:ooxml-core/colors/monochromatic2',
		);
		expect(await zip.file(data!)!.async('string')).toContain('urn:ooxml-core/style/moderate');
		const layout = Object.keys(zip.files).find((name) =>
			/word\/diagrams\/layout.*\.xml$/.test(name),
		);
		expect(await zip.file(layout!)!.async('string')).toContain('layout/default');
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain('<dgm:relIds');

		// Reload the saved file: the diagram comes back with its nodes.
		// Wait for the load: until then the old document's diagram is still there to be clicked.
		await reopenDocument(page, `reloaded-${framework}.docx`, bytes);
		await expect(diagrams(editor)).toHaveCount(1);
		await diagrams(editor)
			.first()
			.click({ position: { x: 4, y: 4 } });
		await openTextPane(editor);
		await expect.poll(() => texts(editor)).toEqual(['Plan', 'Design', 'Build']);
		await expect
			.poll(() => pane(editor).locator('[role="treeitem"]').nth(2).getAttribute('aria-level'))
			.toBe('2');
	});
}

test('editing a Word SmartArt rewrites its data, drops the stale drawing and undoes', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1700, height: 950 });
	await page.goto('/?framework=vanilla');
	const original = await readFile(WORD_FIXTURE);
	await open(page, original, 'smartart-word.docx');
	const editor = page.locator('docx-editor');
	const diagram = diagrams(editor).first();
	await expect(diagram).toHaveAttribute('data-rendering', 'cached-drawing');
	await diagram.click({ position: { x: 4, y: 4 } });
	await openTextPane(editor);
	await expect.poll(() => texts(editor)).toEqual(['Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon']);

	// Add Shape Below the second node from the ribbon, then remove Gamma from the pane.
	await lines(editor).nth(1).click();
	await editor.getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
	const add = editor.getByRole('combobox', { name: 'Add Shape', exact: true });
	await reveal(editor, add);
	await add.selectOption({ label: 'Add Shape Below' });
	await expect
		.poll(() => texts(editor))
		.toEqual(['Alpha', 'Beta', '', 'Gamma', 'Delta', 'Epsilon']);
	await expect(diagram).toHaveAttribute('data-rendering', 'interpreted');
	await lines(editor).nth(2).fill('Child');
	await lines(editor).nth(3).fill('');
	await lines(editor).nth(3).press('Backspace');
	await expect.poll(() => texts(editor)).toEqual(['Alpha', 'Beta', 'Child', 'Delta', 'Epsilon']);

	const { bytes } = await savedDocument(page, 'ui-smartart-word-edited.docx');
	const zip = await JSZip.loadAsync(bytes);
	expect(zip.file('word/diagrams/drawing1.xml')).toBeNull();
	const data = await zip.file('word/diagrams/data1.xml')!.async('string');
	expect(data).toContain('<a:t>Child</a:t>');
	expect(data).not.toContain('<a:t>Gamma</a:t>');
	expect(data).not.toContain('dataModelExt');
	const before = await JSZip.loadAsync(original);
	expect(await zip.file('word/diagrams/layout1.xml')!.async('string')).toBe(
		await before.file('word/diagrams/layout1.xml')!.async('string'),
	);

	// Undo walks the edits back to the original diagram: saving then keeps every diagram part.
	await diagram.click({ position: { x: 4, y: 4 } });
	for (let step = 0; step < 12; step++) await page.keyboard.press('Control+z');
	await expect(diagrams(editor).first()).toHaveAttribute('data-rendering', 'cached-drawing');
	const restored = await JSZip.loadAsync(
		(await savedDocument(page, 'ui-smartart-word-undone.docx')).bytes,
	);
	for (const stem of ['data', 'layout', 'quickStyle', 'colors', 'drawing'])
		expect(await restored.file(`word/diagrams/${stem}1.xml`)!.async('string'), stem).toBe(
			await before.file(`word/diagrams/${stem}1.xml`)!.async('string'),
		);
});
