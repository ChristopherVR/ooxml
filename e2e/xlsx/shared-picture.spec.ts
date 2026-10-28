import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { withDiagramPicture } from '../../src/core/diagram/test-support/picture-drawing';
import { FRAMEWORKS, openLanding, editor } from './helpers';
const bytes = await withDiagramPicture(
	readFileSync(new URL('./support/excel-smartart.xlsx', import.meta.url)),
	'xl',
);
for (const framework of FRAMEWORKS) {
	test(`${framework}: resolves and paints cropped cached SmartArt pictures`, async ({ page }) => {
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'pictures.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(bytes),
		});
		const art = editor(page).locator('office-ui-smartart').first();
		await expect(art.locator('pattern image')).toHaveAttribute('width', '2');
		await expect(art.locator('feFuncA')).toHaveAttribute('slope', '0.5');
		const width = await art.locator('pattern image').evaluate(async (element) => {
			const image = new Image();
			image.src = element.getAttribute('href')!;
			await image.decode();
			return image.naturalWidth;
		});
		expect(width).toBe(2);
		const saved = await editor(page).evaluate(async (element) =>
			Array.from(
				new Uint8Array(
					await (await (element as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
				),
			),
		);
		await page.locator('#landing-file').setInputFiles({
			name: 'saved.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(saved),
		});
		await expect(
			editor(page).locator('office-ui-smartart').first().locator('pattern image'),
		).toHaveAttribute('width', '2');
	});
}
