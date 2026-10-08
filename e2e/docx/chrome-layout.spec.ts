import { expect, test, type Page } from '@playwright/test';
import { openSample } from './helpers';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'];
const editor = (page: Page) => page.locator('docx-editor');
const ribbon = (page: Page) => editor(page).locator('office-ui-ribbon');
const actions = (page: Page) => ribbon(page).locator('office-ui-ribbon-actions');

/**
 * The chrome follows Word 365: one status bar, and the editing mode, Comments and Share at the
 * right end of the ribbon tab row (not in the title bar), in that order, the same shared element
 * Excel uses.
 */
for (const framework of FRAMEWORKS) {
	test(`${framework}: one status bar and Word's tab-row actions`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1440, height: 900 });
		await openSample(page, framework);

		await expect(editor(page).locator('office-ui-status-bar')).toHaveCount(1);
		await expect(editor(page).locator('office-ui-status-bar')).toBeVisible();

		await expect(actions(page)).toHaveCount(1);
		await expect(actions(page)).toHaveAttribute('slot', 'actions');
		const mode = actions(page).getByRole('combobox', { name: 'Editing mode' });
		const comments = actions(page).getByRole('button', { name: 'Comments', exact: true });
		const share = actions(page).getByRole('button', { name: 'Share', exact: true });
		for (const control of [mode, comments, share]) await expect(control).toBeVisible();
		await expect(comments).toHaveText('Comments');
		await expect(share).toHaveText('Share');
		const title = editor(page).locator('office-ui-title-bar');
		await expect(title.getByRole('combobox')).toHaveCount(0);
		await expect(title.locator('office-ui-ribbon-actions')).toHaveCount(0);

		const lastTab = (await ribbon(page).getByRole('tab').last().boundingBox())!;
		const boxes = await Promise.all(
			[mode, comments, share].map(async (control) => (await control.boundingBox())!),
		);
		for (const box of boxes) {
			// On the tab row: vertically centred within the tab's band, right of the tabs.
			const middle = box.y + box.height / 2;
			expect(middle).toBeGreaterThan(lastTab.y);
			expect(middle).toBeLessThan(lastTab.y + lastTab.height);
			expect(box.x).toBeGreaterThan(lastTab.x + lastTab.width);
		}
		// Word's order: the editing mode, then Comments, then Share at the right end.
		expect(boxes[0]!.x).toBeLessThan(boxes[1]!.x);
		expect(boxes[1]!.x).toBeLessThan(boxes[2]!.x);
		const editorBox = (await editor(page).boundingBox())!;
		expect(boxes[2]!.x + boxes[2]!.width).toBeGreaterThan(editorBox.x + editorBox.width - 80);

		// The selector and Comments work from their new place.
		await comments.click();
		await expect(comments).toHaveAttribute('aria-pressed', 'true');
		await expect(editor(page).getByRole('complementary', { name: 'Comments' })).toBeVisible();
		await mode.selectOption('viewing');
		await expect(editor(page)).toHaveAttribute('read-only', '');
		await mode.selectOption('editing');
		await expect(editor(page)).not.toHaveAttribute('read-only', '');
		expect(errors).toEqual([]);
	});
}

test('vanilla: Share starts a session the demo opens in a second window', async ({ page }) => {
	await page.setViewportSize({ width: 1440, height: 900 });
	await openSample(page, 'vanilla');
	const share = actions(page).getByRole('button', { name: 'Share', exact: true });
	const [guest] = await Promise.all([page.waitForEvent('popup'), share.click()]);
	await expect(guest.locator('docx-editor .ProseMirror')).toContainText('Document title');
	await expect(share).toHaveAttribute('aria-pressed', 'true');
	await guest.close();
});
