/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The Office suite's PowerPoint on a phone: the shared Insert sheet and the one-column Format
 * sheet. Runs against a built suite (`bun run build:suite`, then `bun run preview:suite`); set
 * `PPTX_SUITE_URL` (for example http://localhost:8163) to enable it, otherwise it is skipped.
 */
import { test, expect } from '@playwright/test';

const suite = process.env.PPTX_SUITE_URL;

test.skip(!suite, 'set PPTX_SUITE_URL to a served suite-dist');
test.use({ viewport: { width: 412, height: 915 }, hasTouch: true, isMobile: true });

test('Insert opens the phone sheet and a tile inserts through the shared intents', async ({
	page,
}) => {
	await page.goto(`${suite}/apps/powerpoint/`);
	await page.getByText('New presentation').first().click();
	const nav = page.getByRole('navigation', { name: 'Editor actions' });
	await nav.getByRole('button', { name: 'Insert' }).tap();
	for (const name of [
		'Text Box',
		'Shapes',
		'Pictures',
		'Table',
		'Chart',
		'SmartArt',
		'Equation',
		'New Slide',
		'Video',
	]) {
		await expect(page.getByRole('button', { name, exact: true })).toBeVisible();
	}
	await expect(page.getByRole('button', { name: 'Comment', exact: true })).toBeVisible();
	await page.getByRole('button', { name: 'Text Box', exact: true }).tap();
	// The sheet closes once the tile has inserted.
	await expect(page.getByRole('button', { name: 'Text Box', exact: true })).toHaveCount(0);
});

test('the Format sheet is a single column inside the suite', async ({ page }) => {
	await page.goto(`${suite}/apps/powerpoint/`);
	await page.getByText('New presentation').first().click();
	await page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Format' })
		.tap();
	const rects = await page
		.getByRole('tab')
		.evaluateAll((tabs) => tabs.map((tab) => tab.getBoundingClientRect().toJSON()));
	expect(rects).toHaveLength(3);
	expect(new Set(rects.map((rect) => Math.round(rect.top))).size).toBe(1);
	expect(Math.max(...rects.map((rect) => rect.right))).toBeLessThanOrEqual(412);
});
