import { expect, test, type Page } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

async function openOutline(page: Page) {
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
				{ type: 'paragraph', id: 'topic', style: 'Heading1', runs: [{ text: 'Topic' }] },
				{ type: 'paragraph', id: 'child', style: 'Heading2', runs: [{ text: 'Nested' }] },
				{ type: 'paragraph', id: 'body', runs: [{ text: 'Body' }] },
				{ type: 'paragraph', id: 'next', style: 'Heading1', runs: [{ text: 'Next' }] },
			],
		};
	});
	await editor.getByRole('tab', { name: 'View', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Layout view' }).selectOption('outline');
	await expect(editor.locator('.dve-canvas')).toHaveAttribute('data-document-view', 'outline');
	return editor;
}

test('the caret skips hidden paragraphs and a selected collapsed heading acts on its branch', async ({
	page,
}) => {
	const editor = await openOutline(page);
	const topic = editor.getByRole('heading', { name: 'Topic', exact: true });
	const child = editor.locator('[data-id="child"]');
	const next = editor.locator('[data-id="next"]');
	const surface = editor.locator('.dve-paper .ProseMirror');
	await topic.click();
	await editor.getByRole('button', { name: 'Collapse', exact: true }).click();
	await expect(child).toBeHidden();

	await page.keyboard.press('End');
	await page.keyboard.press('Delete');
	await expect(child).toHaveText('Nested');
	await page.keyboard.press('ArrowRight');
	await page.keyboard.press('Backspace');
	await expect(child).toHaveText('Nested');
	await page.keyboard.type('Visible ');
	await expect(next).toHaveText('Visible Next');
	await page.keyboard.press('Control+z');

	await topic.click();
	await editor.getByRole('button', { name: 'Move Down', exact: true }).click();
	await expect(surface.locator(':scope > p').first()).toHaveText('Next');
	await expect(surface.locator(':scope > p').last()).toHaveText('Body');
	await expect(child).toBeHidden();

	await topic.click();
	await page.keyboard.press('Home');
	await page.keyboard.press('Shift+End');
	const copied = () =>
		surface.evaluate((element) => {
			const clipboardData = new DataTransfer();
			element.dispatchEvent(
				new ClipboardEvent('copy', { bubbles: true, cancelable: true, clipboardData }),
			);
			return clipboardData.getData('text/plain');
		});
	await expect.poll(copied).toMatch(/Topic\s+Nested\s+Body/);
	await page.keyboard.press('Delete');
	await expect(surface.locator(':scope > p')).toHaveText(['Next']);
	await page.keyboard.press('Control+z');
	await expect(surface.locator(':scope > p')).toHaveText(['Next', 'Topic', 'Nested', 'Body']);
});

test('rapid Outline commands undo and redo one level change at a time', async ({ page }) => {
	const editor = await openOutline(page);
	const child = editor.getByRole('heading', { name: 'Nested', exact: true });
	await child.click();
	const demote = editor.getByRole('button', { name: 'Demote', exact: true });
	const promote = editor.getByRole('button', { name: 'Promote', exact: true });
	// Both edits occur inside ProseMirror's normal history grouping window.
	await demote.evaluate((button) => {
		if (!(button instanceof HTMLButtonElement)) throw new Error('Expected a demote button');
		button.click();
		button.click();
	});
	await expect(child).toHaveAttribute('aria-level', '4');
	await page.keyboard.press('Control+z');
	await expect(child).toHaveAttribute('aria-level', '3');
	await page.keyboard.press('Control+z');
	await expect(child).toHaveAttribute('aria-level', '2');
	await page.keyboard.press('Control+y');
	await expect(child).toHaveAttribute('aria-level', '3');
	await promote.click();
	await expect(child).toHaveAttribute('aria-level', '2');
	await page.keyboard.press('Control+z');
	await expect(child).toHaveAttribute('aria-level', '3');
});
