import { expect, test, type Locator, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import type { EditorView } from 'prosemirror-view';
import { clickInEditor, reopenDocument } from './helpers';

async function openDraft(page: Page, name: string) {
	const fixture = new URL(
		'../../src/core/docx/layout/fixtures/word-layout-reports/' + name + '.docx',
		import.meta.url,
	);
	await page.setViewportSize({ width: 1400, height: 900 });
	await page.goto('/?framework=vanilla');
	await reopenDocument(page, name + '.docx', await readFile(fixture));
	const editor = page.locator('docx-editor');
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	const view = editor.getByRole('combobox', { name: 'Layout view', exact: true });
	await view.selectOption('draft');
	// Draft is the default view, so the canvas carries no data-document-view until another is chosen.
	await expect(view).toHaveValue('draft');
	await expect(editor.locator('.dve-canvas')).not.toHaveAttribute('data-document-view', /./);
	return editor;
}

const startTag = (editor: Locator, title: string) =>
	editor.locator(`.ProseMirror .dve-sdt-tag[data-tag-title="${title}"]`).first();
const endTag = (paragraph: Locator) => paragraph.locator('.dve-sdt-tag[data-tag-edge="end"]');

async function clickParagraphText(paragraph: Locator) {
	const position = await paragraph.evaluate((p) => {
		const text = document.createTreeWalker(p, NodeFilter.SHOW_TEXT).nextNode();
		if (!text) throw new Error('Expected text inside the paragraph');
		const range = document.createRange();
		range.selectNodeContents(text);
		const textBox = range.getBoundingClientRect();
		const box = p.getBoundingClientRect();
		return {
			x: textBox.left + textBox.width / 2 - box.left,
			y: textBox.top + textBox.height / 2 - box.top,
		};
	});
	await clickInEditor(paragraph, { position });
}

async function expectCaretOnContentSide(tag: Locator) {
	await expect
		.poll(() =>
			tag.evaluate((tag) => {
				const root = tag.getRootNode() as ShadowRoot & { getSelection?: () => Selection | null };
				const { view } = root.host as unknown as { view: EditorView };
				const selection = root.getSelection?.() ?? document.getSelection();
				if (!selection?.isCollapsed || !selection.anchorNode || !(tag instanceof HTMLElement))
					return false;
				const pos = view.posAtDOM(tag, 0);
				if (view.state.selection.head !== pos) return false;

				const caret = document.createRange();
				caret.setStart(selection.anchorNode, selection.anchorOffset);
				caret.collapse(true);
				if (
					!view.dom.contains(caret.startContainer) ||
					view.posAtDOM(caret.startContainer, caret.startOffset) !== pos
				)
					return false;
				const start = tag.dataset.tagEdge === 'start';
				const boundary = document.createRange();
				boundary.selectNode(tag);
				boundary.collapse(!start);
				const order = caret.compareBoundaryPoints(Range.START_TO_START, boundary);
				if (start ? order < 0 : order > 0) return false;
				const rect = caret.getBoundingClientRect();
				const left = rect.height > 0 ? rect.left : view.coordsAtPos(pos, start ? 1 : -1).left;
				const box = tag.getBoundingClientRect();
				return start ? left >= box.right - 1 : left <= box.left + 1;
			}),
		)
		.toBe(true);
}

test('Home and End keep the caret inside paragraph, row and cell tags', async ({ page }) => {
	const editor = await openDraft(page, 'block-content-control-tags');
	const paragraphs = editor.locator('.ProseMirror p');
	for (const [title, first, last] of [
		['Block', /^Lorem ipsum$/, /^amet$/],
		['Rows', /^Consectetur$/, /^eiusmod$/],
		['Cells', /^Consectetur$/, /^adipiscing$/],
	] as const) {
		const start = paragraphs.filter({ hasText: first });
		const end = paragraphs.filter({ hasText: last });
		await clickParagraphText(start);
		await page.keyboard.press('Home');
		await expectCaretOnContentSide(startTag(editor, title));
		await clickParagraphText(end);
		await page.keyboard.press('End');
		await expectCaretOnContentSide(endTag(end));
	}
	const open = startTag(editor, 'Block');
	await clickParagraphText(paragraphs.filter({ hasText: /^Lorem ipsum$/ }));
	await page.keyboard.press('Home');
	await expectCaretOnContentSide(open);
	await page.keyboard.type('Sit ');
	await expect(editor.locator('.ProseMirror p[data-sdt-block]').first()).toHaveText(
		'Sit Lorem ipsum',
	);

	const last = editor.locator('.ProseMirror p[data-sdt-block]').last();
	await clickParagraphText(last);
	await page.keyboard.press('End');
	await expectCaretOnContentSide(endTag(last));
	await page.keyboard.type(' ipsum');
	await expect(last).toHaveText('amet ipsum');
});

test('arrow keys and tag clicks put typed text inside an inline control', async ({ page }) => {
	const editor = await openDraft(page, 'issue-349-content-control-tags');
	const first = editor.locator('.ProseMirror p').first();
	const content = first.locator('.dve-sdt').first();
	const open = startTag(editor, 'Condition');
	const close = endTag(first).first();
	// Tag widgets are re-created when the selection moves, so look the tag up inside one evaluation
	// instead of holding a locator that can resolve to a node replaced before it is read.
	const position = (selector: string) =>
		first.evaluate((paragraph, selector) => {
			const root = paragraph.getRootNode() as ShadowRoot;
			const { view } = root.host as unknown as { view: EditorView };
			const tag = paragraph.querySelector(selector);
			if (!tag) throw new Error('Missing tag ' + selector);
			return { caret: view.state.selection.head, tag: view.posAtDOM(tag, 0) };
		}, selector);
	const openSelector = '.dve-sdt-tag[data-tag-title="Condition"]';
	const closeSelector = '.dve-sdt-tag[data-tag-edge="end"]';

	await clickInEditor(content);
	const start = await position(openSelector);
	expect(start.caret).toBeGreaterThan(start.tag);
	for (let pos = start.caret; pos > start.tag; pos--) {
		await page.keyboard.press('ArrowLeft');
		await expect.poll(async () => (await position(openSelector)).caret).toBe(pos - 1);
	}
	await expectCaretOnContentSide(open);
	await page.keyboard.type('Lorem ');
	await expect(content).toHaveText('Lorem alpha');

	await clickInEditor(content);
	const end = await position(closeSelector);
	expect(end.caret).toBeLessThan(end.tag);
	for (let pos = end.caret; pos < end.tag; pos++) {
		await page.keyboard.press('ArrowRight');
		await expect.poll(async () => (await position(closeSelector)).caret).toBe(pos + 1);
	}
	await expectCaretOnContentSide(close);
	await page.keyboard.press('ArrowRight');
	await expect.poll(async () => (await position(closeSelector)).caret).toBe(end.tag + 1);
	await close.click();
	await expectCaretOnContentSide(close);
	await page.keyboard.type(' ipsum');
	await expect(content).toHaveText('Lorem alpha ipsum');
});
