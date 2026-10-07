import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { createWorkbook, createEditSession, saveXlsx } from 'ooxml-core/xlsx';
import { buildChartGradientDef, resolveChartGradient } from 'ooxml-core/chart';
import { parseXml, NS } from 'ooxml-core/xml';
import { parseDrawingFill, resolveDrawingColor } from 'ooxml-core/diagram';
import paths from '../../src/core/chart/__fixtures__/native-gradient-path-profiles.json' with { type: 'json' };
import circleShape from '../../src/core/chart/__fixtures__/native-gradient-circle-shape-profiles.json' with { type: 'json' };
import native from '../../src/core/chart/__fixtures__/native-gradient-raster.json' with { type: 'json' };
import profiles from '../../src/core/chart/__fixtures__/native-gradient-linear-profiles.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

const samples = [
	...native.cases.map((sample) => ({
		...sample,
		paintStops: 256,
		samples: sample.samples.map((pixel) => ({ ...pixel, alpha: 255 })),
	})),
	...profiles.cases.map((sample) => ({
		...sample,
		paintStops: sample.profile.startsWith('three') ? 3 : 2,
	})),
];

const runs = [
	{
		name: 'native circle and rectangular-shape chart gradient raster',
		gap: false,
		cases: circleShape.cases.map((sample) => ({
			...sample,
			paintStops: sample.profile.includes('transparent') ? 2 : 256,
		})),
	},
	{
		name: 'native rectangular chart gradient raster',
		gap: false,
		cases: paths.cases.map((sample) => ({ ...sample, paintStops: 256 })),
	},
	{
		name: 'native scaled chart gradient raster',
		gap: false,
		cases: samples.filter(
			(sample) =>
				!sample.name.startsWith('three-transparent') &&
				sample.name !== 'coincident-angle-135-480x300',
		),
	},
	{
		name: 'known native translucent three-stop mismatch',
		gap: true,
		cases: samples.filter((sample) => sample.name === 'three-transparent-angle-135-300x300'),
	},
	{
		name: 'known native coincident edge mismatch',
		gap: true,
		cases: samples.filter((sample) => sample.name === 'coincident-angle-135-480x300'),
	},
];
for (const run of runs)
	for (const framework of FRAMEWORKS)
		test(`${run.name} in ${framework}`, async ({ page }) => {
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
			for (const sample of run.cases) {
				const zip = await JSZip.loadAsync(base);
				const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
				zip.file(
					'xl/drawings/drawing1.xml',
					drawing.replace(
						'cx="4572000" cy="2857500"',
						`cx="${(sample.width / 2) * 9525}" cy="${(sample.height / 2) * 9525}"`,
					),
				);
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
				const nativeFill = parseDrawingFill(
					parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
				)!;
				if (nativeFill.kind !== 'gradient') throw new Error('Expected native gradient');
				const expected = buildChartGradientDef(
					'native',
					resolveChartGradient(nativeFill, (color) => resolveDrawingColor(color)),
					{ width: sample.width, height: sample.height, shape: 'rect' },
				);
				const gradient = editor(page)
					.locator(`${expected.kind === 'rectPath' ? 'pattern' : expected.kind}[id$="-chartArea"]`)
					.first();
				if (expected.kind === 'rectPath') {
					await expect(gradient.locator('image')).toHaveAttribute('href', expected.href);
				} else if (expected.kind === 'linearGradient') {
					for (const attr of ['x1', 'y1', 'x2', 'y2'] as const)
						await expect(gradient).toHaveAttribute(attr, String(expected[attr]));
					await expect(gradient.locator('stop')).toHaveCount(sample.paintStops);
				} else if (expected.kind === 'radialGradient') {
					for (const attr of ['cx', 'cy', 'r'] as const)
						await expect(gradient).toHaveAttribute(attr, String(expected[attr]));
					if (expected.gradientTransform)
						await expect(gradient).toHaveAttribute('gradientTransform', expected.gradientTransform);
					else await expect(gradient).not.toHaveAttribute('gradientTransform');
					await expect(gradient.locator('stop')).toHaveCount(sample.paintStops);
				} else throw new Error('Unexpected native paint');
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
						Array.from(ctx.getImageData(pixel.x, pixel.y, 1, 1).data),
					);
				}, sample);
				if (run.gap)
					test.fail(
						true,
						'The measured native raster mismatch remains open; strict pixel comparison is intentionally expected to fail.',
					);
				for (const [index, pixel] of sample.samples.entries()) {
					expect(
						Math.abs(pixels[index]![3]! - pixel.alpha),
						`${framework} ${sample.name} alpha`,
					).toBeLessThanOrEqual(2);
					for (const channel of [0, 1, 2])
						expect(
							Math.abs(
								(pixels[index]![channel]! * pixels[index]![3]!) / 255 -
									(pixel.rgb[channel]! * pixel.alpha) / 255,
							),
							`${framework} ${sample.name} at ${pixel.x},${pixel.y} channel ${channel}`,
						).toBeLessThanOrEqual(pixel.alpha === 255 ? 2 : 3);
				}
			}
			expect(errors).toEqual([]);
		});
