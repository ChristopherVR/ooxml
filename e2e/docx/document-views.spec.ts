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
	const body = canvas.locator('.ProseMirror');
	await body.press('Control+Home');
	await canvas.evaluate((element) => (element.scrollTop = 200));
	await setView(editor, 'read');
	await expect(canvas).toHaveAttribute('data-document-view', 'read');
	await expect(body).toHaveAttribute('contenteditable', 'false');
	await expect(body).toHaveCSS('column-count', '2');
	const controls = editor.getByRole('group', { name: 'Read Mode' });
	const status = controls.getByRole('status');
	await expect(status).toHaveText(/Page 1 of [2-9]/);
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
	const first = editor.getByRole('heading', { level: 1, name: 'First branch' });
	const child = editor.locator('[data-id="child"]');
	await first.click();
	await editor.getByRole('button', { name: 'Collapse', exact: true }).click();
	await expect(child).toBeHidden();
	await editor.getByRole('button', { name: 'Expand', exact: true }).click();
	await expect(child).toBeVisible();
	const level = editor.getByRole('combobox', { name: 'Show Level' });
	await level.selectOption('1');
	await expect(child).toBeHidden();
	await level.selectOption('10');
	await expect(child).toBeVisible();
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
	const saved = await editor.evaluate(async (element) => [
		...(await (element as DocxEditorElement).saveBytes()),
	]);
	const reopened = (await loadDocx(Uint8Array.from(saved))).model;
	expect(reopened.blocks[1]).toMatchObject({ type: 'paragraph', outlineLevel: 1 });
});
