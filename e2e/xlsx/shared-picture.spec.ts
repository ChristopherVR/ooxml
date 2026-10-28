import { createEditSession, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
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

for (const framework of FRAMEWORKS) {
	test(`${framework}: paints cropped and corrected Excel pictures after save and reload`, async ({
		page,
	}) => {
		const workbook = await loadXlsx(
			readFileSync(
				new URL('../../src/core/xlsx/__fixtures__/excel-features.xlsx', import.meta.url),
			),
		);
		const sheet = workbook.sheets.findIndex((value) =>
			value.drawings.some((drawing) => drawing.kind === 'image'),
		);
		const index = workbook.sheets[sheet]!.drawings.findIndex((drawing) => drawing.kind === 'image');
		const session = createEditSession(workbook);
		session.setDrawingAnchor(sheet, index, {
			from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
			ext: { cx: 952500, cy: 476250 },
		});
		session.setPictureCrop(sheet, index, { cropLeft: 0.25, cropRight: 0.25 });
		session.editPictureEffects(sheet, index, {
			grayscale: true,
			brightness: 0.2,
			contrast: -0.1,
			alpha: 0.5,
		});
		workbook.activeSheet = sheet;
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'corrected.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await saveXlsx(workbook)),
		});
		const picture = editor(page)
			.locator('.xg-obj')
			.filter({ has: page.locator('img') })
			.first();
		await expect(picture.locator('img')).toHaveCSS('width', '200px');
		await expect(picture.locator('img')).toHaveCSS('left', '-50px');
		await expect(picture.locator('feFuncA')).toHaveAttribute('slope', '0.5');
		await expect(picture.locator('feColorMatrix')).toHaveAttribute('values', '0');
		const saved = await editor(page).evaluate(async (element) =>
			Array.from(
				new Uint8Array(
					await (await (element as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
				),
			),
		);
		const reopened = await loadXlsx(Buffer.from(saved));
		expect(reopened.sheets[sheet]!.drawings[index]).toMatchObject({
			kind: 'image',
			crop: { cropLeft: 0.25, cropRight: 0.25 },
		});
	});
}
