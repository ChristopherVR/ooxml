import { expect, test, type Locator } from '@playwright/test';
import { loadDocx } from '../../viewers/docx/packages/core/src/index';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

async function setView(editor: Locator, mode: string) {
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view' }).selectOption(mode);
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: switch to Web Layout and focus without losing the caret`, async ({
		page,
	}) => {
		await page.setViewportSize({ width: 1280, height: 800 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element) => {
			const host = element as DocxEditorElement;
			const model = host.documentModel;
			if (!model) throw new Error('Expected a document');
			host.documentModel = {
				...model,
				pageColor: 'EFEBE2',
				blocks: [
					{
						type: 'paragraph',
						id: 'a',
						runs: [{ text: 'Before' }, { text: '', break: 'page' }, { text: 'after' }],
					},
				],
			};
		});
		const canvas = editor.locator('.dve-canvas');
		const paper = canvas.locator('.dve-paper');
		await setView(editor, 'web');
		await expect(canvas).toHaveAttribute('data-document-view', 'web');
		await expect(paper).toHaveCSS('background-color', 'rgb(239, 235, 226)');
		await expect(paper.locator('.dve-break-marker')).toBeHidden();
		expect(await paper.evaluate((el) => el.clientWidth)).toBe(
			await canvas.evaluate((el) => el.clientWidth),
		);
		const body = paper.locator('.ProseMirror');
		await body.press('Control+End');
		await editor.getByRole('button', { name: 'Focus', exact: true }).click();
		await expect(editor.locator('.dve-ribbon')).toBeHidden();
		await expect(editor.getByRole('button', { name: 'Exit Focus' })).toBeVisible();
		await expect(editor.locator('.dve-frame')).toHaveCSS('position', 'fixed');
		await page.keyboard.type(' focused');
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-ribbon')).toBeVisible();
		await expect(canvas).toHaveAttribute('data-document-view', 'web');
		await page.keyboard.type(' restored');
		await expect(body).toContainText('after focused restored');
	});
}

test('Read Mode flows the document across columns and returns to the editing position', async ({
	page,
}) => {
	await page.setViewportSize({ width: 1280, height: 640 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel;
		if (!model) throw new Error('Expected a document');
		host.documentModel = {
			...model,
			blocks: Array.from({ length: 36 }, (_, index) => ({
				type: 'paragraph',
				id: `reading-${index}`,
				runs: [
					{
						text:
							`Paragraph ${index + 1}. ` + 'The document continues in reading order. '.repeat(8),
					},
				],
			})),
		};
	});
	const canvas = editor.locator('.dve-canvas');
	const paper = canvas.locator('.dve-paper');
	const body = canvas.locator('.ProseMirror');
	await body.press('Control+Home');
	await canvas.evaluate((element) => (element.scrollTop = 200));
	await setView(editor, 'read');
	await expect(canvas).toHaveAttribute('data-document-view', 'read');
	await expect(body).toHaveAttribute('contenteditable', 'false');
	await expect(body).toHaveCSS('column-count', '2');
	const [canvasBox, paperBox] = await Promise.all([canvas.boundingBox(), paper.boundingBox()]);
	if (!canvasBox || !paperBox) throw new Error('Expected a reading viewport');
	expect(paperBox.width).toBeLessThan(canvasBox.width);
	expect(paperBox.x + paperBox.width / 2).toBeCloseTo(canvasBox.x + canvasBox.width / 2, 0);
	const controls = editor.getByRole('group', { name: 'Read Mode' });
	const status = controls.getByRole('status');
	await expect(status).toHaveText(/^Page 1 of (?:[2-9]|\d{2,})$/);
	const frame = editor.locator('.dve-frame');
	await frame.evaluate((element) => {
		const fields = document.createElement('div');
		fields.dataset.readerTestFields = '';
		fields.innerHTML = `<input value="Ab"><textarea>Ab</textarea>
			<select><option>One</option><option>Two</option></select>
			<div contenteditable="true">Ab</div><div role="dialog" tabindex="0">Dialog</div>`;
		element.append(fields);
	});
	const fields = frame.locator('[data-reader-test-fields]');
	for (const selector of ['input', 'textarea', '[contenteditable]']) {
		const field = fields.locator(selector);
		await field.press('End');
		await field.press('ArrowLeft');
		await field.press('Backspace');
		await field.press('Space');
		if (selector === '[contenteditable]') await expect(field).toHaveText('b');
		else await expect(field).toHaveValue(' b');
	}
	await fields.locator('select').press('End');
	await expect(fields.locator('select')).toHaveValue('Two');
	await fields.getByRole('dialog').press('ArrowRight');
	await expect(status).toHaveText(/Page 1 of/);
	await fields.evaluate((element) => element.remove());
	await canvas.focus();
	await page.keyboard.type('EditingBlocked');
	await expect(body).not.toContainText('EditingBlocked');
	await page.keyboard.press('ArrowRight');
	await expect(status).toHaveText(/Page 2 of/);
	await expect
		.poll(() => canvas.locator('.dve-paper').evaluate((el) => el.scrollLeft))
		.toBeGreaterThan(0);
	await controls.getByRole('button', { name: 'Previous page' }).click();
	await expect(status).toHaveText(/Page 1 of/);
	await page.setViewportSize({ width: 600, height: 640 });
	await expect(body).toHaveCSS('column-count', '1');
	await expect(status).toHaveText(/^Page 1 of (?:[2-9]|\d{2,})$/);
	await page.setViewportSize({ width: 1280, height: 640 });
	await page.keyboard.press('Escape');
	await expect(canvas).toHaveAttribute('data-document-view', 'draft');
	await expect.poll(() => canvas.evaluate((el) => el.scrollTop)).toBe(200);
	await expect(body).toHaveAttribute('contenteditable', 'true');
	await page.keyboard.type('Start. ');
	await expect(body.locator('p').first()).toContainText('Start. Paragraph 1.');
});

test('closing Outline restores Print Layout', async ({ page }) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const canvas = editor.locator('.dve-canvas');
	await setView(editor, 'print');
	await expect(canvas).toHaveAttribute('data-document-view', 'print');
	await setView(editor, 'outline');
	await editor.getByRole('button', { name: 'Close Outline View', exact: true }).click();
	await expect(canvas).toHaveAttribute('data-document-view', 'print');
});

test('a following table stays within the Outline editing area, including in Focus', async ({
	page,
}) => {
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel;
		if (!model) throw new Error('Expected a document');
		host.documentModel = {
			...model,
			blocks: [
				{ type: 'paragraph', id: 'topic', outlineLevel: 2, runs: [{ text: 'Topic' }] },
				{ type: 'paragraph', id: 'body', runs: [{ text: 'Body' }] },
				{
					type: 'table',
					id: 'table',
					rows: [
						[
							{
								paragraphs: [{ type: 'paragraph', id: 'cell', runs: [{ text: 'A table cell' }] }],
							},
						],
					],
				},
			],
		};
	});
	await setView(editor, 'outline');
	await editor.getByRole('heading', { name: 'Topic' }).click();
	await editor.getByRole('button', { name: 'Demote', exact: true }).click();
	await expect(editor.getByRole('heading', { name: 'Topic' })).toHaveAttribute('aria-level', '3');
	const body = editor.locator('.dve-paper .ProseMirror');
	const table = body.getByRole('table');
	async function checkTableBounds() {
		const [area, bounds] = await Promise.all([body.boundingBox(), table.boundingBox()]);
		if (!area || !bounds) throw new Error('Expected a table in the editing area');
		expect(bounds.x).toBeCloseTo(area.x, 0);
		expect(bounds.x + bounds.width).toBeLessThanOrEqual(area.x + area.width + 1);
	}
	await checkTableBounds();
	await editor.getByRole('button', { name: 'Focus', exact: true }).click();
	await expect(editor.getByRole('button', { name: 'Exit Focus' })).toBeVisible();
	await checkTableBounds();
});

test('Outlining filters and collapses branches and saves an undoable edit', async ({ page }) => {
	await page.setViewportSize({ width: 1280, height: 800 });
	await page.goto('/?framework=vanilla');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	await editor.evaluate((element) => {
		const host = element as DocxEditorElement;
		const model = host.documentModel;
		if (!model) throw new Error('Expected a document');
		host.documentModel = {
			...model,
			blocks: [
				{ type: 'paragraph', id: 'one', style: 'Heading1', runs: [{ text: 'First branch' }] },
				{ type: 'paragraph', id: 'child', outlineLevel: 2, runs: [{ text: 'Nested topic' }] },
				{ type: 'paragraph', id: 'body', runs: [{ text: 'Body text '.repeat(30), bold: true }] },
				{ type: 'paragraph', id: 'two', outlineLevel: 1, runs: [{ text: 'Second branch' }] },
			],
		};
	});
	await setView(editor, 'outline');
	await expect(editor.getByRole('tab', { name: 'Outlining' })).toHaveAttribute(
		'aria-selected',
		'true',
	);
	const first = editor.getByRole('heading', { name: 'First branch' });
	const child = editor.locator('[data-id="child"]');
	const paper = editor.locator('.dve-paper');
	await expect(paper).toHaveCSS('box-shadow', 'none');
	await expect(paper).toHaveCSS('background-color', 'rgb(255, 255, 255)');
	const group = editor.getByRole('group', { name: 'Outline Tools controls' });
	await expect(
		group.getByRole('button', { name: 'Promote', exact: true }).locator('svg'),
	).toBeVisible();
	await expect(
		editor
			.getByRole('group', { name: 'Close controls' })
			.getByRole('button', { name: 'Close Outline View' }),
	).toHaveClass(/ribbon-large/);
	await expect(first.locator('.dve-outline-marker')).toHaveAttribute('data-marker', 'minus');
	await expect(editor.locator('[data-id="body"] .dve-outline-marker')).toHaveAttribute(
		'data-marker',
		'bullet',
	);
	await first.click();
	await editor.getByRole('button', { name: 'Collapse', exact: true }).click();
	await expect(child).toBeHidden();
	await expect(first.locator('.dve-outline-marker')).toHaveAttribute('data-marker', 'plus');
	await editor.getByRole('button', { name: 'Expand', exact: true }).click();
	await expect(child).toBeVisible();
	const level = editor.getByRole('combobox', { name: 'Show Level' });
	await level.selectOption('1');
	await expect(child).toBeHidden();
	await level.selectOption('10');
	await expect(child).toBeVisible();
	await child.click();
	const paragraphLevel = group.getByRole('combobox', { name: 'Outline Level', exact: true });
	await expect(paragraphLevel).toHaveValue('2');
	await paragraphLevel.selectOption('3');
	await expect(child).toHaveAttribute('aria-level', '3');
	await paragraphLevel.selectOption('2');
	await editor.getByRole('button', { name: 'Show First Line Only' }).click();
	await expect(editor.locator('.dve-canvas')).toHaveAttribute('data-outline-first-line', '');
	await editor.getByRole('button', { name: 'Show Text Formatting' }).click();
	await expect(editor.locator('.dve-canvas')).toHaveAttribute('data-outline-plain', '');
	await first.click();
	await editor.getByRole('button', { name: 'Move Down', exact: true }).click();
	await expect(editor.locator('.dve-paper .ProseMirror > p').first()).toHaveText('Second branch');
	await page.keyboard.press('Control+z');
	await expect(editor.locator('.dve-paper .ProseMirror > p').first()).toHaveText('First branch');
	await child.click();
	await editor.getByRole('button', { name: 'Promote', exact: true }).click();
	await expect(child).toHaveAttribute('aria-level', '1');
	await first.click();
	await editor.getByRole('button', { name: 'Demote', exact: true }).click();
	await expect(first).toHaveAttribute('aria-level', '2');
	const saved = await editor.evaluate(async (element) => [
		...(await (element as DocxEditorElement).saveBytes()),
	]);
	const reopened = (await loadDocx(Uint8Array.from(saved))).model;
	expect(reopened.blocks[0]).toMatchObject({ type: 'paragraph', style: 'Heading2' });
	expect(reopened.blocks[0]).not.toHaveProperty('outlineLevel');
	expect(reopened.blocks[1]).toMatchObject({ type: 'paragraph', outlineLevel: 1 });
});
