import { expect, test } from '@playwright/test';
import { clippedRibbonLabels } from '../ribbon-labels';
import { newDocument, reveal } from './helpers';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: Insert groups fit and commands use the correct tabs`, async ({ page }) => {
		await page.setViewportSize({ width: 1600, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const tab = (name: string) => editor.getByRole('tab', { name, exact: true });
		await tab('Insert').click();
		const insert = editor.locator('.ribbon-panel[data-panel="Insert"]');
		const group = (name: string) => insert.locator(`.ribbon-group[data-label="${name}"]`);
		await expect(insert.locator(':scope > .ribbon-group')).toHaveCount(10);
		expect(
			await insert
				.locator(':scope > .ribbon-group')
				.evaluateAll((groups) => groups.map((group) => group.getAttribute('data-label'))),
		).toEqual([
			'Pages',
			'Tables',
			'Illustrations',
			'Add-ins',
			'Media',
			'Links',
			'Comments',
			'Header & Footer',
			'Text',
			'Symbols',
		]);
		await expect(group('Add-ins').getByRole('button', { name: 'Get Add-ins' })).toBeVisible();
		await expect.poll(() => insert.evaluate(clippedRibbonLabels)).toEqual([]);
		await page.setViewportSize({ width: 1280, height: 900 });
		await expect.poll(() => insert.evaluate(clippedRibbonLabels)).toEqual([]);
		await expect(
			insert.locator('[aria-label="Link to Previous"], [aria-label="Insert column break"]'),
		).toHaveCount(0);
		await expect(group('Symbols').locator('select[aria-label="Symbol"]')).toHaveCount(1);
		const video = group('Media').locator('button[aria-label="Online Video"]');
		await expect(video).toBeDisabled();
		await expect(video).toHaveAttribute(
			'aria-description',
			'Inserting online video is not supported',
		);

		await editor.locator('.ProseMirror').click();
		await page.keyboard.type('Review me');
		await page.keyboard.press('Home');
		await page.keyboard.press('ArrowRight');
		const comment = editor.locator('button[aria-label="Add comment"]').filter({ visible: true });
		const foldedComments = group('Comments').locator('.ribbon-overflow-button');
		if (await foldedComments.isVisible()) await foldedComments.click();
		await comment.click();
		const pane = editor.locator('.dve-comments-panel');
		await expect(pane).toBeVisible();
		await expect(pane.getByRole('textbox', { name: 'New comment' })).toBeFocused();
		await page.keyboard.type('Inserted from the ribbon');
		await pane.getByRole('button', { name: 'Add comment', exact: true }).click();
		await expect(pane.locator('.dve-comment-thread')).toContainText('Inserted from the ribbon');
		await expect(editor.locator('.ProseMirror .dve-comment-range')).toHaveText('Review');
		await pane.getByRole('button', { name: 'Close comments' }).click();

		await tab('Layout').click();
		const column = editor.getByRole('button', { name: 'Insert column break', exact: true });
		await expect(
			editor.locator(
				'.ribbon-panel[data-panel="Layout"] .ribbon-group[data-label="Breaks"] button[aria-label="Insert column break"]',
			),
		).toHaveCount(1);
		await reveal(editor, column);
		await column.click();
		await expect(editor.locator('.ProseMirror [data-docx-break="column"]')).toHaveCount(1);

		await tab('Insert').click();
		const foldedHeader = group('Header & Footer').locator('.ribbon-overflow-button');
		if (await foldedHeader.isVisible()) await foldedHeader.click();
		await editor.getByRole('combobox', { name: 'Header', exact: true }).selectOption('blank');
		await editor.locator('.dve-header [data-slot="default"]').dblclick();
		await expect(tab('Header & Footer')).toHaveAttribute('aria-selected', 'true');
		await expect(
			editor.locator('.ribbon-panel[data-panel="Header & Footer"] [aria-label="Link to Previous"]'),
		).toHaveCount(1);
	});
}

for (const [locale, label] of [
	['en', 'Header & Footer'],
	['fr', 'En-tête et pied de page'],
	['de', 'Kopf- und Fußzeile'],
	['es', 'Encabezado y pie de página'],
	['zh-CN', '页眉和页脚'],
] as const) {
	test(`Header & Footer fits at both ribbon widths in ${locale}`, async ({ page }) => {
		await page.goto('/');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.evaluate((element, locale) => {
			if ('locale' in element) element.locale = locale;
		}, locale);
		await editor.locator('[role="tab"][data-tab="insert"]').click();
		const insert = editor.locator('.ribbon-panel[data-panel="Insert"]');
		const header = insert.locator('[data-label="Header & Footer"]');
		for (const width of [1600, 1280]) {
			await page.setViewportSize({ width, height: 900 });
			await expect(header).toHaveAttribute('data-caption', label);
			await expect
				.poll(() =>
					header.evaluate((group, label) => {
						const caption = group.querySelector('.ribbon-overflow-button > span');
						if (!(caption instanceof HTMLElement) || !caption.getClientRects().length) {
							const style = getComputedStyle(group, '::after');
							const context = document.createElement('canvas').getContext('2d');
							if (!context || style.content !== JSON.stringify(label)) return false;
							context.font = `${style.fontSize} ${style.fontFamily}`;
							return context.measureText(label).width <= group.clientWidth + 1;
						}
						const text = caption.firstChild;
						if (text?.textContent !== label) return false;
						const box = caption.getBoundingClientRect();
						const range = document.createRange();
						const lines = new Set<number>();
						for (const word of text.textContent.matchAll(/\S+/g)) {
							range.setStart(text, word.index);
							range.setEnd(text, word.index + word[0].length);
							const rects = range.getClientRects();
							if (rects.length !== 1) return false;
							for (const rect of rects) {
								if (rect.left < box.left - 1 || rect.right > box.right + 1) return false;
								lines.add(rect.top);
							}
						}
						return lines.size <= 2 && caption.scrollHeight <= caption.clientHeight + 1;
					}, label),
				)
				.toBe(true);
		}
	});
}

test('Review adds an empty comment anchor between words', async ({ page }) => {
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror');
	await body.click();
	await page.keyboard.type('Alpha  beta');
	await page.keyboard.press('Home');
	for (let offset = 0; offset < 6; offset++) await page.keyboard.press('ArrowRight');
	await editor.getByRole('tab', { name: 'Review', exact: true }).click();
	const comment = editor.getByRole('button', { name: 'Add comment', exact: true });
	await reveal(editor, comment);
	await comment.click();
	const pane = editor.locator('.dve-comments-panel');
	await expect(pane.getByRole('textbox', { name: 'New comment' })).toBeFocused();
	await page.keyboard.type('Empty anchor');
	await pane.getByRole('button', { name: 'Add comment', exact: true }).click();
	await expect(pane.locator('.dve-comment-thread')).toContainText('Empty anchor');
	await expect(body.locator('[data-docx-comment-point]')).toHaveCount(1);
	await expect(body.locator('.dve-comment-range')).toHaveCount(0);
	expect(await body.textContent()).toBe('Alpha  beta');
});
