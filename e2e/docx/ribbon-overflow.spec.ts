import { expect, test } from '@playwright/test';
import { dialogByHeading, newDocument, openSample, reveal, saveButton } from './helpers';

test.describe('ribbon overflow', () => {
	test.beforeEach(async ({ page }) => {
		await page.setViewportSize({ width: 1000, height: 700 });
		await openSample(page);
	});

	test('a narrow ribbon folds groups into dropdowns that still run their commands', async ({
		page,
	}) => {
		const editor = page.locator('docx-editor');
		const panel = editor.locator('.ribbon-panel[data-panel="Home"]');
		const fits = () => panel.evaluate((element) => element.scrollWidth <= element.clientWidth + 1);
		expect(await fits()).toBe(true);
		const folded = editor.locator('.ribbon-group[data-collapsed] > .ribbon-overflow-button');
		await expect(folded.first()).toBeVisible();
		await editor.locator('.ProseMirror p').first().click();
		await folded.filter({ hasText: 'Paragraph' }).click();
		await editor.getByRole('button', { name: 'Align center' }).click();
		await expect(editor.locator('.ProseMirror p').first()).toHaveCSS('text-align', 'center');
		// A command ends the dropdown, and the controls return to their group.
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(0);
	});

	test('widening the window unfolds the groups again', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ribbon-group[data-collapsed]').first()).toBeAttached();
		await page.setViewportSize({ width: 1700, height: 700 });
		await expect(editor.locator('.ribbon-group[data-collapsed]')).toHaveCount(0);
		await expect(editor.getByRole('button', { name: 'Align center' })).toBeVisible();
	});

	test('Escape closes an open group dropdown', async ({ page }) => {
		const editor = page.locator('docx-editor');
		await editor.locator('.ribbon-group[data-collapsed] > .ribbon-overflow-button').first().click();
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(1);
		await page.keyboard.press('Escape');
		await expect(editor.locator('.ribbon-overflow-panel')).toHaveCount(0);
	});
});

test.describe('on a phone', () => {
	test.use({ viewport: { width: 320, height: 720 }, hasTouch: true, isMobile: true });

	test('scrolls wide ribbon groups to reach SmartArt, then undoes, redoes and saves', async ({
		page,
	}) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'Insert', exact: true }).tap();
		await editor.getByRole('button', { name: 'Illustrations', exact: true }).tap();
		const group = editor.locator('.ribbon-overflow-panel');
		await expect(group).toBeInViewport({ ratio: 1 });
		await group.getByRole('button', { name: 'More', exact: true }).tap();
		await expect.poll(() => group.evaluate((panel) => panel.scrollLeft)).toBeGreaterThan(0);
		const smartArt = group.getByRole('button', { name: 'Insert SmartArt', exact: true });
		await smartArt.tap();
		const picker = dialogByHeading(editor, 'Choose a SmartArt Graphic');
		await picker.getByRole('option', { name: 'Basic Block List', exact: true }).tap();
		await picker.getByRole('button', { name: 'OK', exact: true }).tap();

		const diagram = editor.locator('.ProseMirror office-ui-smartart');
		await expect(diagram).toHaveCount(1);
		await editor.locator('.qat').getByRole('button', { name: 'Undo', exact: true }).tap();
		await expect(diagram).toHaveCount(0);
		await editor.locator('.qat').getByRole('button', { name: 'Redo', exact: true }).tap();
		await expect(diagram).toHaveCount(1);
		const download = page.waitForEvent('download');
		await saveButton(page).tap();
		expect((await download).suggestedFilename()).toMatch(/\.docx$/);
	});

	test('applies paragraph formatting, counts words and sets the view zoom', async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		const editor = page.locator('docx-editor');
		const body = editor.locator('.ProseMirror');
		await body.tap();
		await page.keyboard.type('three short words');
		await editor.getByRole('button', { name: 'Paragraph', exact: true }).tap();
		await editor.getByRole('button', { name: 'Align center', exact: true }).tap();
		await expect(body.locator('p')).toHaveCSS('text-align', 'center');

		await editor.getByRole('tab', { name: 'Review', exact: true }).tap();
		const wordCount = editor.getByRole('button', { name: 'Word count', exact: true });
		await reveal(editor, wordCount);
		await wordCount.tap();
		await expect(editor.locator('.dve-word-count')).toContainText('3');
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-word-count')).toHaveCount(0);

		await editor.getByRole('tab', { name: 'View', exact: true }).tap();
		const zoom = editor.getByRole('button', { name: 'Zoom dialog', exact: true });
		await reveal(editor, zoom);
		await zoom.tap();
		const dialog = dialogByHeading(editor, 'Zoom');
		await dialog.getByLabel('Percent', { exact: true }).fill('75');
		await dialog.getByRole('button', { name: 'OK', exact: true }).tap();
		await expect(editor.locator('.dve-status')).toContainText('75%');
	});
});
