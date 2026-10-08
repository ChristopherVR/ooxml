import { expect, test } from '@playwright/test';
import {
	createWorkbook,
	createEditSession,
	saveXlsx,
	chartView,
	renderChartSvg,
	loadXlsx,
} from 'ooxml-core/xlsx';
import { parseXml, NS } from 'ooxml-core/xml';
import { parseDrawingFill } from 'ooxml-core/diagram';
import native from '../../src/core/chart/__fixtures__/native-gradient-series-path-profiles.json' with { type: 'json' };
import { FRAMEWORKS, editor, openLanding, pageErrors } from './helpers';

// Open browser-only gaps (core models pass both at pixel centres and corners):
// - circle-transparent at 391,479: browser 3.04 premultiplied levels (limit 3); the core
//   radial model is 2.18 off (native green premul 133.2, model 135.4), and the native red
//   channel reads 254 from export truncation, so browser gradient quantization adds ~0.9.
// - centre shape at 302,376: browser 3 levels (limit 2); the core band image gives 23-24
//   against native 23, but the green ramp there is ~3 levels per pixel, so a sub-pixel
//   registration difference of the per-mark pattern explains it.
// Needed evidence: the browser's per-pixel values for both marks with the rendered mark
// bounds, or native captures at a larger chart size where the ramp is shallower.
const known = [
	'series-column-path-corner-circle-transparent-angle-2-300x300',
	'series-column-path-center-shape-angle-1-300x300',
];
const runs = [
	{
		name: 'native bar and column path raster',
		gap: false,
		cases: native.cases.filter((sample) => !known.includes(sample.name)),
	},
	...known.map((name) => ({
		name: `known native series raster mismatch ${name}`,
		gap: true,
		cases: native.cases.filter((sample) => sample.name === name),
	})),
];
for (const run of runs)
	for (const framework of FRAMEWORKS)
		test(`${run.name} in ${framework}`, async ({ page }) => {
			const errors = pageErrors(page);
			const mismatches: Array<{ name: string; x: number; y: number; error: number }> = [];
			await openLanding(page, framework);
			for (const sample of run.cases) {
				const book = createWorkbook(),
					session = createEditSession(book);
				const fill = parseDrawingFill(
					parseXml(`<a:spPr xmlns:a="${NS.a}">${sample.fillXml}</a:spPr>`).documentElement,
				)!;
				session.addChart(0, {
					chartType: sample.target === 'series-bar' ? 'bar' : 'column',
					showLegend: false,
					series: [{ name: 'Value', categories: ['A'], values: [1], fill }],
					formatting: {
						sourceXml: '',
						entries: {
							valueAxis: { sourceXml: '', labelsVisible: false },
							categoryAxis: { sourceXml: '', labelsVisible: false },
						},
					},
					anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
				});
				const loaded = await loadXlsx(await saveXlsx(book));
				const chart = loaded.sheets[0]!.drawings[0]!;
				if (chart.kind !== 'chart') throw new Error('Expected chart');
				const view = chartView(loaded, 0, chart, () => []);
				// Match the native mark's physical aspect. This isolates fill fidelity from the
				// still-independent native chart layout; it does not claim that layout matches.
				const markSize = (w: number, h: number) => {
					const doc = parseXml(renderChartSvg(view, w, h));
					const mark = [...doc.getElementsByTagName('g')]
						.find((g) => g.getAttribute('data-chart-series') === '0')!
						.getElementsByTagName('rect')[0]!;
					return [Number(mark.getAttribute('width')), Number(mark.getAttribute('height'))];
				};
				const base = markSize(600, 600),
					dx = markSize(6600, 600),
					dy = markSize(600, 6600);
				const slopeX = (dx[0]! - base[0]!) / 6000,
					slopeY = (dy[1]! - base[1]!) / 6000;
				const aspect = sample.paintBounds.width / sample.paintBounds.height;
				let width = 600,
					height = 600,
					error = Infinity;
				// The viewer rounds the outer SVG size. Choose integer bounds with the closest
				// mark aspect rather than letting that rounding change a native radial profile.
				for (let w = 200; w < 2400; w++) {
					const markWidth = base[0]! + (w - 600) * slopeX;
					const h = Math.round(600 + (markWidth / aspect - base[1]!) / slopeY);
					if (h < 200 || h > 2400) continue;
					const difference = Math.abs(markWidth / (base[1]! + (h - 600) * slopeY) - aspect);
					if (difference < error) {
						width = w;
						height = h;
						error = difference;
					}
				}
				chart.anchor.ext = { cx: width * 9525, cy: height * 9525 };
				await page.locator('#landing-file').setInputFiles({
					name: `${sample.name}.xlsx`,
					mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
					buffer: Buffer.from(await saveXlsx(loaded)),
				});
				await expect(editor(page).getByText(`${sample.name}.xlsx`, { exact: true })).toBeVisible();
				const mark = editor(page)
					.locator('g[data-chart-series="0"][data-chart-point="0"] rect')
					.first();
				await expect(mark).toBeAttached();
				const size = await mark.evaluate((rect) => [
					Number(rect.getAttribute('width')),
					Number(rect.getAttribute('height')),
				]);
				expect(size[0]! / size[1]!, `${sample.name} physical mark aspect`).toBeCloseTo(aspect, 4);
				const pixels = await mark.evaluate(async (rect, sample) => {
					const id = rect.getAttribute('fill')!.slice(5, -1);
					const def = (rect as SVGElement).ownerSVGElement!.querySelector(`[id="${id}"]`)!;
					if (!def.id.includes('-mark0-s0')) throw new Error('Expected per-mark gradient');
					const bounds = sample.paintBounds;
					const image = new Image();
					image.src = `data:image/svg+xml,${encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="${bounds.width}" height="${bounds.height}"><defs>${def.outerHTML}</defs><rect width="100%" height="100%" fill="url(#${id})"/></svg>`)}`;
					await image.decode();
					const canvas = document.createElement('canvas');
					canvas.width = bounds.width;
					canvas.height = bounds.height;
					const ctx = canvas.getContext('2d')!;
					ctx.drawImage(image, 0, 0);
					return sample.samples.map((pixel) =>
						Array.from(ctx.getImageData(pixel.x - bounds.x, pixel.y - bounds.y, 1, 1).data),
					);
				}, sample);
				for (const [index, pixel] of sample.samples.entries()) {
					expect(
						Math.abs(pixels[index]![3]! - pixel.alpha),
						`${sample.name} alpha`,
					).toBeLessThanOrEqual(2);
					for (const channel of [0, 1, 2]) {
						const error = Math.abs(
							(pixels[index]![channel]! * pixels[index]![3]!) / 255 -
								(pixel.rgb[channel]! * pixel.alpha) / 255,
						);
						if (error > (pixel.alpha === 255 ? 2 : 3))
							mismatches.push({ name: sample.name, x: pixel.x, y: pixel.y, error });
					}
				}
			}
			expect(errors).toEqual([]);
			if (run.gap)
				test.fail(
					true,
					'The measured native sampling discrepancy remains open; tolerances are unchanged.',
				);
			expect(mismatches).toEqual([]);
		});
