import { expect, test, type Page } from '@playwright/test';
import { FRAMEWORKS, editor, openSample, pageErrors, part, ribbon, statusBar } from './helpers';

/** Every leaf text in the editor that reads exactly `text`, light and shadow trees alike. */
const leafTexts = (page: Page, text: string) =>
	editor(page).evaluate((host, wanted) => {
		let count = 0;
		const walk = (root: ParentNode) => {
			for (const node of root.querySelectorAll('*')) {
				if (!node.childElementCount && node.textContent?.trim() === wanted) count += 1;
				if (node.shadowRoot) walk(node.shadowRoot);
			}
		};
		walk(host.shadowRoot!);
		return count;
	}, text);

/**
 * The chrome follows Excel 365: one status bar reading Ready once, and the editing mode, Comments
 * and Share at the right end of the ribbon tab row (not in the title bar), in that order.
 */
for (const framework of FRAMEWORKS) {
	test(`${framework}: one status bar and Excel's tab-row actions`, async ({ page }) => {
		const errors = pageErrors(page);
		await page.setViewportSize({ width: 1440, height: 900 });
		await openSample(page, framework);

		await expect(editor(page).locator('office-ui-status-bar')).toHaveCount(1);
		await expect(statusBar(page)).toHaveCount(1);
		await expect(statusBar(page)).toBeVisible();
		await expect.poll(() => leafTexts(page, 'Ready')).toBe(1);

		const mode = ribbon(page).getByRole('combobox', { name: 'Editing mode' });
		const comments = ribbon(page).getByRole('button', { name: 'Comments', exact: true });
		const share = ribbon(page).locator('office-ui-ribbon-actions [part="share"]');
		for (const control of [mode, comments, share]) await expect(control).toBeVisible();
		await expect(share).toHaveText('Share');
		await expect(part(page, 'title-bar').getByRole('combobox')).toHaveCount(0);
		await expect(
			part(page, 'title-bar').locator('office-ui-ribbon-actions [part="share"]'),
		).toHaveCount(0);

		const tabs = ribbon(page).getByRole('tab');
		const lastTab = (await tabs.last().boundingBox())!;
		const boxes = await Promise.all(
			[mode, comments, share].map(async (c) => (await c.boundingBox())!),
		);
		for (const box of boxes) {
			// On the tab row: vertically centred within the tab's band...
			const middle = box.y + box.height / 2;
			expect(middle).toBeGreaterThan(lastTab.y);
			expect(middle).toBeLessThan(lastTab.y + lastTab.height);
			// ...and right of the tabs.
			expect(box.x).toBeGreaterThan(lastTab.x + lastTab.width);
		}
		// Excel's order: the editing mode, then Comments, then Share at the right end.
		expect(boxes[0]!.x).toBeLessThan(boxes[1]!.x);
		expect(boxes[1]!.x).toBeLessThan(boxes[2]!.x);
		const editorBox = (await editor(page).boundingBox())!;
		expect(boxes[2]!.x + boxes[2]!.width).toBeGreaterThan(editorBox.x + editorBox.width - 80);
		expect(errors).toEqual([]);
	});
}
