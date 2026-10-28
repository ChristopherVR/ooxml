import { loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { FRAMEWORKS, openLanding, editor } from './helpers';
const workbook = await loadXlsx(
	readFileSync(new URL('../../src/core/xlsx/__fixtures__/excel-features.xlsx', import.meta.url)),
);
const sheet = workbook.sheets.findIndex((s) => s.drawings.some((d) => d.kind === 'image'));
const image = workbook.sheets[sheet]!.drawings.find((d) => d.kind === 'image')!;
if (image.kind !== 'image') throw new Error('Missing image');
image.layout = { tile: { scaleX: 2, scaleY: 2, flip: 'xy', alignment: 'ctr' } };
image.effects = [
	{
		kind: 'standard',
		name: 'duotone',
		colors: [
			{ kind: 'srgb', value: '000000', transforms: [] },
			{ kind: 'srgb', value: 'FF0000', transforms: [] },
		],
	},
];
workbook.activeSheet = sheet;
const bytes = Buffer.from(await saveXlsx(workbook));
for (const framework of FRAMEWORKS)
	test(`${framework}: tiled rich Excel picture survives browser save`, async ({ page }) => {
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'tile.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: bytes,
		});
		const picture = editor(page)
			.locator('.xg-obj svg')
			.filter({ has: page.locator('pattern') })
			.first();
		await expect(picture.locator('pattern[data-ready="true"]')).toHaveCount(1);
		await expect(picture.locator('pattern image')).toHaveCount(4);
		await expect(picture.locator('feComponentTransfer')).toHaveCount(1);
		const saved = await editor(page).evaluate(async (el) =>
			Array.from(
				new Uint8Array(
					await (await (el as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
				),
			),
		);
		const reopened = await loadXlsx(Buffer.from(saved));
		expect(reopened.sheets[sheet]!.drawings.find((d) => d.kind === 'image')).toMatchObject({
			layout: { tile: { flip: 'xy' } },
			effects: [{ kind: 'standard', name: 'duotone' }],
		});
		await page.screenshot({ path: `test-results/picture-rich-excel-${framework}.png` });
	});
for (const framework of FRAMEWORKS)
	test(`${framework}: applies nonlinear Excel picture stages in source order`, async ({ page }) => {
		const mixed = await loadXlsx(bytes);
		const image = mixed.sheets[sheet]!.drawings.find((d) => d.kind === 'image')!;
		if (image.kind !== 'image') throw new Error('Missing image');
		image.effects = [
			{
				kind: 'standard',
				name: 'clrRepl',
				colors: [{ kind: 'srgb', value: 'FF0000', transforms: [] }],
			},
			{ kind: 'standard', name: 'hsl', values: { hue: 60, saturation: 1, luminance: 1 } },
			{ kind: 'alphaModFix', amount: 0.5 },
		];
		await openLanding(page, framework);
		await page.locator('#landing-file').setInputFiles({
			name: 'mixed.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await saveXlsx(mixed)),
		});
		const picture = editor(page)
			.locator('.xg-obj svg')
			.filter({ has: page.locator('pattern') })
			.first();
		await expect(picture.locator('pattern')).toHaveAttribute('data-pixel-effects-ready', 'true');
		const rgba = await picture
			.locator('pattern image')
			.first()
			.evaluate(async (element) => {
				const image = new Image();
				image.src = element.getAttribute('href')!;
				await image.decode();
				const canvas = document.createElement('canvas');
				canvas.width = image.naturalWidth;
				canvas.height = image.naturalHeight;
				const context = canvas.getContext('2d')!;
				context.drawImage(image, 0, 0);
				return Array.from(context.getImageData(0, 0, 1, 1).data);
			});
		expect(rgba.slice(0, 3)).toEqual([255, 255, 0]);
		expect(rgba[3]).toBeGreaterThanOrEqual(127);
		expect(rgba[3]).toBeLessThanOrEqual(128);
	});
