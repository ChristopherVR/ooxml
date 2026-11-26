/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright API */
import { expect, test } from '@playwright/test';

import { loadDeck, ribbonTab } from './support/deck';

for (const { width, embedded } of [
	{ width: 800, embedded: false },
	{ width: 1400, embedded: true },
]) {
	test(`Ribbon tabs stay on one line at ${embedded ? '700px embedded' : '800px'}`, async ({
		page,
	}) => {
		await page.setViewportSize({ width, height: 900 });
		await loadDeck(page);
		const tabs = page.locator('pptx-ui-ribbon-tabs:visible');
		if (embedded) {
			await tabs.evaluate((element) => {
				const editor = element.closest<HTMLElement>('[data-pptx-editor-chrome]');
				if (!editor) throw new Error('Editor root is missing');
				editor.style.width = '700px';
			});
		}
		await expect
			.poll(() =>
				tabs.evaluate((element) => {
					const row = element.shadowRoot?.querySelector('.head');
					const height = Number.parseFloat(
						getComputedStyle(element).getPropertyValue('--office-ribbon-tab-height'),
					);
					return Math.abs((row?.getBoundingClientRect().height ?? 0) - height);
				}),
			)
			.toBeLessThanOrEqual(0.5);
		await expect(tabs.getByRole('tab', { name: 'Home', exact: true })).toHaveCSS(
			'padding-inline-start',
			'5px',
		);
		await expect
			.poll(() =>
				page.evaluate(
					() => document.documentElement.scrollWidth - document.documentElement.clientWidth,
				),
			)
			.toBeLessThanOrEqual(1);
	});
}

for (const width of [800, 1280]) {
	test(`Insert fits and opens a docked comment draft at ${width}px`, async ({ page }) => {
		await page.setViewportSize({ width, height: 900 });
		await loadDeck(page);
		await ribbonTab(page, 'Insert').click();
		const row = page.locator('[data-pptx-chrome="ribbon-content"]:visible').first();
		await expect(row.locator('[data-ribbon-group="insert.comments"]')).toBeVisible();
		await expect
			.poll(() => row.evaluate((element) => element.scrollWidth - element.clientWidth))
			.toBeLessThanOrEqual(1);
		const group = row.locator('[data-ribbon-group="insert.comments"]');
		const command = group.locator('[data-ribbon-control="insert.comments.comment"]');
		const face = group.locator('.face');
		const opener = (await face.isVisible()) ? face : command.getByRole('button');
		if (await face.isVisible()) {
			await face.click();
			const faceBox = await face.boundingBox();
			const popupBox = await group.locator('.row').boundingBox();
			expect(faceBox).not.toBeNull();
			expect(popupBox).not.toBeNull();
			const rightGap = width - (popupBox!.x + popupBox!.width);
			expect(popupBox!.x).toBeGreaterThanOrEqual(0);
			expect(rightGap).toBeGreaterThanOrEqual(0);
			expect(Math.abs(popupBox!.x - faceBox!.x) <= 4 || rightGap <= 10).toBe(true);
			expect(popupBox!.y).toBeGreaterThanOrEqual(faceBox!.y + faceBox!.height - 1);
		}
		await command.getByRole('button').click();
		await expect(row.locator('[data-ribbon-group][data-open]')).toHaveCount(0);
		const pane = page.locator('[data-pptx-comments-panel]:visible');
		const draft = pane.locator(
			'textarea[data-pptx-comment-draft], [data-pptx-comment-draft] textarea',
		);
		await expect(pane).toHaveCount(1);
		await expect(draft).toBeFocused();
		const ribbon = page.locator('[data-pptx-chrome="ribbon"]:visible').first();
		const status = page.locator('pptx-ui-status-bar:visible').first();
		const paneBox = await pane.boundingBox();
		const ribbonBox = await ribbon.boundingBox();
		expect(paneBox!.y).toBeGreaterThanOrEqual(ribbonBox!.y + ribbonBox!.height - 1);
		if (await status.isVisible())
			expect(paneBox!.y + paneBox!.height).toBeLessThanOrEqual((await status.boundingBox())!.y + 1);
		await draft.fill('Draft stays with the slide');
		await draft.press('Escape');
		await expect(pane).toHaveCount(0);
		await expect(opener).toBeFocused();
		const header = page
			.locator('pptx-ui-ribbon-actions')
			.getByRole('button', { name: 'Comments', exact: true });
		await expect(header).toHaveAttribute('aria-pressed', 'false');
		await header.click();
		await expect(header).toHaveAttribute('aria-pressed', 'true');
		await expect(draft).toHaveValue('Draft stays with the slide');
		await pane.getByRole('button', { name: 'Add Comment', exact: true }).click();
		await expect(pane.locator('[data-comment-id]').first()).toBeVisible();
		await header.click();
		await expect(pane).toHaveCount(0);
		await expect(header).toHaveAttribute('aria-pressed', 'false');
	});
}

test('Home uses Drawing Arrange and Review shares the Comments task pane', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 900 });
	await loadDeck(page);
	await ribbonTab(page, 'Home').click();
	await expect(page.locator('[data-ribbon-group="home.arrange"]:visible')).toHaveCount(0);
	const drawing = page.locator('[data-ribbon-group="home.drawing"]:visible');
	await expect(drawing.getByRole('button', { name: 'Quick Styles' })).toBeVisible();
	await drawing.getByRole('button', { name: 'Arrange', exact: true }).click();
	await expect(drawing.getByRole('menuitem', { name: 'Selection Pane' })).toBeVisible();
	await page.keyboard.press('Escape');

	await ribbonTab(page, 'Review').click();
	const showComments = page
		.locator('[data-ribbon-control="review.comments.showComments"]')
		.getByRole('button');
	await showComments.click();
	await expect(showComments).toHaveAttribute('aria-pressed', 'true');
	const pane = page.locator('[data-pptx-comments-panel]:visible');
	await expect(pane).toHaveCount(1);
	await pane.getByRole('button', { name: 'Close', exact: true }).click();
	await expect(showComments).toHaveAttribute('aria-pressed', 'false');
	await expect(showComments).toBeFocused();
	await page
		.locator('[data-ribbon-control="review.comments.newComment"]')
		.getByRole('button')
		.click();
	await expect(
		pane.locator('textarea[data-pptx-comment-draft], [data-pptx-comment-draft] textarea'),
	).toBeFocused();
	await showComments.click();
	await expect(pane).toHaveCount(0);

	const header = page
		.locator('pptx-ui-ribbon-actions')
		.getByRole('button', { name: 'Comments', exact: true });
	await expect(header).toHaveAttribute('aria-pressed', 'false');
	await header.click();
	await expect(header).toHaveAttribute('aria-pressed', 'true');
	await pane.getByRole('button', { name: 'Close', exact: true }).click();
	await expect(header).toHaveAttribute('aria-pressed', 'false');
	await expect(header).toBeFocused();
});

test('Insert Illustrations labels Freeform Shape and Curve', async ({ page }) => {
	await page.setViewportSize({ width: 1920, height: 900 });
	await loadDeck(page);
	await ribbonTab(page, 'Insert').click();
	const illustrations = page.locator('[data-ribbon-group="insert.illustrations"]');
	for (const [id, label] of [
		['freeformShape', 'Freeform: Shape'],
		['curve', 'Curve'],
	]) {
		const tool = illustrations.locator(`[data-pptx-drawing-tool="${id}"]`);
		const button = tool.getByRole('button', { name: label, exact: true });
		await expect(button).toBeVisible();
		await expect(button).toHaveText(label);
		await expect(tool).toHaveAttribute('title', label);
	}
});
