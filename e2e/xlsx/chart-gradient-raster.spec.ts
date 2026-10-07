import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, saveXlsx } from 'ooxml-core/xlsx';
import { buildChartGradientDef } from 'ooxml-core/chart';
import native from '../../src/core/chart/__fixtures__/native-gradient-raster.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`native scaled chart gradient raster in ${framework}`, async ({ page }) => {
		const errors = pageErrors(page);
		const book = createWorkbook();
		createEditSession(book).addChart(0, {
			chartType: 'column',
			series: [{ name: 'Value', categories: ['A'], values: [1] }],
			anchor: {
				from: { row: 1, col: 1, rowOffset: 0, colOffset: 0 },
				ext: { cx: 4572000, cy: 2857500 },
			},
		});
		const base = await saveXlsx(book);
		await openLanding(page, framework);
		for (const sample of native.cases) {
			const zip = await JSZip.loadAsync(base);
			const xml = await zip.file('xl/charts/chart1.xml')!.async('string');
			zip.file(
				'xl/charts/chart1.xml',
				xml.replace('</c:chartSpace>', `<c:spPr>${sample.fillXml}</c:spPr></c:chartSpace>`),
			);
			await page.locator('#landing-file').setInputFiles({
				name: `${sample.name}.xlsx`,
				mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				buffer: await zip.generateAsync({ type: 'nodebuffer' }),
			});
			const gradient = editor(page).locator('linearGradient[id$="-chartArea"]').first();
			const expected = buildChartGradientDef('native', {
				type: 'linear',
				angle: sample.angle,
				scaled: true,
				stops: [],
			});
			if (expected.kind !== 'linearGradient') throw new Error('Expected linear gradient');
			for (const attr of ['x1', 'y1', 'x2', 'y2'] as const)
				await expect(gradient).toHaveAttribute(attr, String(expected[attr]));
			await expect(gradient.locator('stop')).toHaveCount(256);
			const pixels = await gradient.evaluate(async (element, sample) => {
				const image = new Image();
				const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${sample.width}" height="${sample.height}"><defs>${element.outerHTML}</defs><rect width="100%" height="100%" fill="url(#${element.id})"/></svg>`;
				image.src = `data:image/svg+xml,${encodeURIComponent(svg)}`;
				await image.decode();
				const canvas = document.createElement('canvas');
				canvas.width = sample.width;
				canvas.height = sample.height;
				const ctx = canvas.getContext('2d')!;
				ctx.drawImage(image, 0, 0);
				return sample.samples.map((pixel) =>
					Array.from(ctx.getImageData(pixel.x, pixel.y, 1, 1).data).slice(0, 3),
				);
			}, sample);
			for (const [index, pixel] of sample.samples.entries())
				for (const channel of [0, 1, 2])
					expect(
						Math.abs(pixels[index]![channel]! - pixel.rgb[channel]!),
						`${framework} ${sample.name} at ${pixel.x},${pixel.y} channel ${channel}`,
					).toBeLessThanOrEqual(2);
		}
		expect(errors).toEqual([]);
	});
