/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * The Office suite's PowerPoint on a phone: the shared Insert sheet and the one-column Format
 * sheet. Runs against a built suite (`bun run build:suite`, then `bun run preview:suite`); set
 * `PPTX_SUITE_URL` (for example http://localhost:8163) to enable it, otherwise it is skipped.
 */
import { test, expect, type Page } from '@playwright/test';

const suite = process.env.PPTX_SUITE_URL;

test.skip(!suite, 'set PPTX_SUITE_URL to a served suite-dist');
test.use({ viewport: { width: 412, height: 915 }, hasTouch: true, isMobile: true });

async function openPhonePresentation(page: Page): Promise<void> {
	await page.goto(`${suite}/apps/powerpoint/`);
	await page.getByText('New presentation').first().tap();
	await expect(page.locator('#doc-position')).toHaveText('Slide 1 / 1', { timeout: 30_000 });
	await page.locator('#layout-toggle').click();
	await expect(page.locator('body')).toHaveAttribute('data-layout', 'office');
}

test('Insert opens the phone sheet and a tile inserts through the shared intents', async ({
	page,
}) => {
	await openPhonePresentation(page);
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
	await openPhonePresentation(page);
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

test('the suite phone Insert sheet commits WordArt and Icons', async ({ page }) => {
	await openPhonePresentation(page);
	const insert = page
		.getByRole('navigation', { name: 'Editor actions' })
		.getByRole('button', { name: 'Insert', exact: true });
	const sheet = page.locator('pptx-ui-mobile-insert');
	const stage = page.locator('[data-pptx-viewport]');
	const picker = page.locator('office-ui-dialog[open]');
	await insert.tap();
	await sheet.getByRole('button', { name: 'WordArt', exact: true }).tap();
	await picker.getByRole('textbox', { name: 'WordArt text' }).fill('Phone presentation');
	await picker.locator('.pptx-wordart-grid button').first().tap();
	await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect(sheet).toBeHidden();
	await expect(stage.getByText('Phone presentation', { exact: true })).toBeVisible();

	await insert.tap();
	await sheet.getByRole('button', { name: 'Icons', exact: true }).tap();
	await picker.getByRole('searchbox', { name: 'Search icons' }).fill('Light bulb');
	await picker.getByRole('button', { name: 'Light bulb', exact: true }).tap();
	await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
	await expect(sheet).toBeHidden();
	await expect(stage.getByRole('img', { name: 'Light bulb', exact: true })).toBeVisible();
});
