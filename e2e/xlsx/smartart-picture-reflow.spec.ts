import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { withNativePictureLayout } from '../../src/core/diagram/test-support/picture-layout';
import { FRAMEWORKS, editor, openLanding } from './helpers';

for (const framework of FRAMEWORKS) {
	test(`${framework}: native picture SmartArt reflows and saves its images`, async ({ page }) => {
		await page.setViewportSize({ width: 1700, height: 950 });
		await openLanding(page, framework);
		const bytes = await withNativePictureLayout(
			readFileSync(
				new URL('../../src/core/xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url),
			),
			'xl',
		);
		await page
			.locator('input[type=file]')
			.first()
			.setInputFiles({
				name: 'pictures.xlsx',
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: Buffer.from(bytes),
			});
		const art = editor(page).locator('office-ui-smartart').first();
		await expect(art.locator('image')).toHaveCount(3);
		await editor(page)
			.locator('.xg-obj:has(office-ui-smartart)')
			.first()
			.click({ position: { x: 20, y: 20 } });
		await editor(page).getByRole('tab', { name: 'SmartArt Design', exact: true }).click();
		await editor(page).getByRole('button', { name: 'Text Pane', exact: true }).first().click();
		await editor(page)
			.locator('office-ui-smartart-text-pane input[data-node]')
			.first()
			.fill('Picture reflow caption');
		await expect(art.locator('image')).toHaveCount(3);
		await editor(page).evaluate(async (node) => {
			const host = node as unknown as {
				save(): Promise<Blob>;
				load(bytes: Uint8Array): Promise<void>;
			};
			await host.load(new Uint8Array(await (await host.save()).arrayBuffer()));
		});
		await expect(art.locator('image')).toHaveCount(3);
		await expect(art).toContainText('Picture reflow caption');
	});
}
