import { expect, test } from '@playwright/test';
import { clippedRibbonLabels } from '../ribbon-labels';
import { newDocument } from './helpers';

// Word breaks "Get Add-ins" as "Get" / "Add-ins": a large caption breaks only at its spaces.
for (const framework of ['vanilla', 'react']) {
	test(`${framework}: large Insert captions never break after a hyphen`, async ({ page }) => {
		await page.setViewportSize({ width: 1920, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		const insert = editor.locator('.ribbon-panel[data-panel="Insert"]');
		const caption = insert.getByRole('button', { name: 'Get Add-ins' }).locator(':scope > span');
		await expect(caption).toHaveText('Get Add-ins');
		const lines = await caption.evaluate((span) =>
			[...span.querySelectorAll('.ribbon-caption-word')].map((word) => ({
				word: word.textContent,
				boxes: word.getClientRects().length,
				top: Math.round(word.getBoundingClientRect().top),
			})),
		);
		expect(lines.map((line) => line.word)).toEqual(['Get', 'Add-ins']);
		// Each word is one unbroken box, and the hyphenated word is not split across lines.
		expect(lines.every((line) => line.boxes === 1)).toBe(true);
		await expect.poll(() => insert.evaluate(clippedRibbonLabels)).toEqual([]);
	});
}
