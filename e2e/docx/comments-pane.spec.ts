import { expect, test, type Locator, type Page } from '@playwright/test';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { newDocument } from './helpers';

// Word's comments pane is the shared `office-ui-comments-pane`; its controls live in the
// element's open shadow root, which Playwright's locators pierce. The pane only fires events, so
// each step also checks the document model the editor keeps.
const comments = (page: Page) =>
	page.locator('docx-editor').evaluate((element) =>
		((element as DocxEditorElement).documentModel?.comments ?? []).map((comment) => ({
			text: comment.text,
			parentId: comment.parentId ?? null,
			resolved: Boolean(comment.resolved),
		})),
	);

/** Selects `word` inside the text node that holds it (comment anchors split the paragraph). */
async function selectWord(surface: Locator, word: string) {
	await surface.evaluate((root, target) => {
		const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
		let text = walker.nextNode();
		while (text && !text.textContent!.includes(target)) text = walker.nextNode();
		if (!text) throw new Error(`No text node holds ${target}`);
		const at = text.textContent!.indexOf(target);
		window.getSelection()!.setBaseAndExtent(text, at, text, at + target.length);
	}, word);
	await expect
		.poll(() => surface.page().evaluate(() => window.getSelection()?.toString()))
		.toBe(word);
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: the comments pane adds, replies, resolves, reopens, navigates and deletes`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1600, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const surface = editor.locator('.ProseMirror');
		await surface.click();
		await page.keyboard.type('Alpha beta gamma');

		await selectWord(surface, 'Alpha');
		await editor.locator('office-ui-ribbon-actions [part="comments"]').click();
		const pane = editor.getByRole('complementary', { name: 'Comments' });
		await expect(pane).toBeVisible();
		const box = pane.getByRole('textbox', { name: 'New comment', exact: true });
		const add = pane.getByRole('button', { name: 'Add comment', exact: true });
		await box.fill('First note');
		await add.click();
		const cards = pane.locator('.dve-comment-card');
		await expect(cards).toHaveCount(1);

		await selectWord(surface, 'gamma');
		await box.fill('Second note');
		await add.click();
		await expect(cards).toHaveCount(2);
		const first = cards.filter({ hasText: 'First note' });
		const second = cards.filter({ hasText: 'Second note' });
		await expect
			.poll(() => comments(page))
			.toEqual([
				{ text: 'First note', parentId: null, resolved: false },
				{ text: 'Second note', parentId: null, resolved: false },
			]);

		// Reply to the first thread from its own reply box.
		await first.getByRole('textbox', { name: 'Reply', exact: true }).fill('A reply');
		await first.getByRole('button', { name: 'Reply', exact: true }).click();
		await expect(first.locator('.dve-comment-thread')).toHaveCount(2);
		await expect(first).toContainText('A reply');
		await expect
			.poll(async () => (await comments(page)).filter((comment) => comment.parentId).length)
			.toBe(1);

		// Resolve, then reopen.
		const resolved = async () =>
			(await comments(page)).find((comment) => comment.text === 'First note')?.resolved;
		await first.getByRole('button', { name: 'Resolve', exact: true }).click();
		await expect(first).toHaveAttribute('data-resolved', 'true');
		await expect.poll(resolved).toBe(true);
		await first.getByRole('button', { name: 'Reopen', exact: true }).click();
		await expect(first).toHaveAttribute('data-resolved', 'false');
		await expect.poll(resolved).toBe(false);

		// Arrow keys, Home and End move between threads; Enter selects one and the caret goes to
		// its anchor, which makes it the active thread.
		await first.focus();
		await expect(first).toBeFocused();
		await page.keyboard.press('ArrowDown');
		await expect(second).toBeFocused();
		await page.keyboard.press('Home');
		await expect(first).toBeFocused();
		await page.keyboard.press('End');
		await expect(second).toBeFocused();
		await page.keyboard.press('Enter');
		await expect(second).toHaveAttribute('aria-current', 'true');
		await expect(first).not.toHaveAttribute('aria-current', 'true');
		// Escape closes the pane and returns focus to the document, at that anchor.
		await page.keyboard.press('Escape');
		await expect(pane).toBeHidden();
		await expect(surface).toBeFocused();
		await page.keyboard.type('X');
		await expect(surface.locator('p').first()).toHaveText('Alpha beta Xgamma');
		await page.keyboard.press('Control+z');
		await expect(surface.locator('p').first()).toHaveText('Alpha beta gamma');

		// Clicking a thread also goes to its anchor.
		await editor.locator('office-ui-ribbon-actions [part="comments"]').click();
		await expect(pane).toBeVisible();
		await first.locator('.author').first().click();
		await expect(first).toHaveAttribute('aria-current', 'true');

		// Delete the second thread from its root comment.
		await second
			.locator('.dve-comment-thread')
			.first()
			.getByRole('button', { name: 'Delete', exact: true })
			.click();
		await expect(cards).toHaveCount(1);
		await expect(pane).not.toContainText('Second note');
		await expect
			.poll(async () => (await comments(page)).map((comment) => comment.text))
			.toEqual(['First note', 'A reply']);
		expect(errors).toEqual([]);
	});
}
