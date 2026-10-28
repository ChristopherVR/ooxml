import { pictureGridPng } from '../../src/core/diagram/test-support/picture-grid';
import JSZip from 'jszip';
import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileInput } from './helpers';
const zip = await JSZip.loadAsync(
	readFileSync(
		new URL(
			'../../src/core/docx/__fixtures__/review-inline/picture-insert-accepted.docx',
			import.meta.url,
		),
	),
);
for (const name of Object.keys(zip.files))
	if (name.startsWith('word/media/') && name.endsWith('.png')) zip.file(name, pictureGridPng());
const xml = (await zip.file('word/document.xml')!.async('string'))
	.replace(/<wp:extent\b[^>]*\/>/, '<wp:extent cx="1828800" cy="1219200"/>')
	.replace(/<a:ext\b[^>]*\/>/, '<a:ext cx="1828800" cy="1219200"/>');
zip.file(
	'word/document.xml',
	xml
		.replace(
			/<a:stretch>[\s\S]*?<\/a:stretch>/,
			'<a:tile sx="200000" sy="200000" tx="95250" algn="ctr" flip="xy"/>',
		)
		.replace(
			/<a:blip([^>]*?)(\/>|>)/,
			'<a:blip$1><a:duotone><a:srgbClr val="000000"/><a:srgbClr val="FF0000"/></a:duotone></a:blip>',
		),
);
const bytes = Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: tiled rich Word picture survives browser save`, async ({ page }) => {
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'tile.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: bytes,
		});
		const picture = page.locator('docx-editor .ProseMirror .dve-picture svg').first();
		await expect(picture.locator('pattern[data-ready="true"]')).toHaveCount(1);
		await expect(picture.locator('pattern image')).toHaveCount(4);
		await expect(picture.locator('feComponentTransfer')).toHaveCount(1);
		expect(await picture.evaluate((el) => el.getBoundingClientRect().width)).toBeGreaterThan(20);
		const saved = await page
			.locator('docx-editor')
			.evaluate(async (el) =>
				Array.from(await (el as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
			);
		const archive = await JSZip.loadAsync(Buffer.from(saved));
		const output = await archive.file('word/document.xml')!.async('string');
		expect(output).toContain('flip="xy"');
		expect(output).toContain('duotone');
		await page.screenshot({ path: `test-results/picture-rich-word-${framework}.png` });
	});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: applies nonlinear picture stages in source order`, async ({ page }) => {
		const mixed = await JSZip.loadAsync(bytes);
		const xml = await mixed.file('word/document.xml')!.async('string');
		mixed.file(
			'word/document.xml',
			xml.replace(
				/<a:duotone>[\s\S]*?<\/a:duotone>/,
				'<a:clrRepl><a:srgbClr val="FF0000"/></a:clrRepl><a:hsl hue="3600000" sat="100000" lum="100000"/><a:alphaModFix amt="50000"/>',
			),
		);
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'mixed.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await mixed.generateAsync({ type: 'uint8array' })),
		});
		const picture = page.locator('docx-editor .ProseMirror .dve-picture svg').first();
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
				return Array.from(context.getImageData(5, 5, 1, 1).data);
			});
		expect(rgba.slice(0, 3)).toEqual([255, 255, 0]);
		expect(rgba[3]).toBeGreaterThanOrEqual(127);
		expect(rgba[3]).toBeLessThanOrEqual(128);
		await expect(picture.locator('pattern image').first()).not.toHaveAttribute('filter');
	});
