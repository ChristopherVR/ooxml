import { test, expect } from '@playwright/test';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import type { VisioDocument } from 'ooxml-core/visio';

const directory = process.env.VISIO_NATIVE_GRADIENT_RASTER_DIR;
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	for (const starOnly of [false, true]) {
		test(`${framework}: ${starOnly ? 'records unresolved star gradient fidelity' : 'measures saved gradient interiors against native PNG'}`, async ({
			page,
		}) => {
			test.skip(!directory, 'Set VISIO_NATIVE_GRADIENT_RASTER_DIR to the native raster capture.');
			const evidence = JSON.parse(
				(await readFile(join(directory!, 'evidence.json'), 'utf8')).replace(/^\uFEFF/, ''),
			) as {
				cases: {
					name: string;
					direction: number;
					stopCount: number;
					alpha: boolean;
					shapeId: string;
					kind:
						| 'rectangle'
						| 'ellipse'
						| 'triangle'
						| 'notched'
						| 'pentagon'
						| 'chevron'
						| 'ushape'
						| 'star';
					outline?: [number, number][];
				}[];
			};
			const samples = await Promise.all(
				evidence.cases
					.filter((item) => item.direction <= 13 && (item.kind === 'star') === starOnly)
					.map(async (item) => {
						const source = await readFile(join(directory!, item.name + '.svg'), 'utf8');
						const frame = /viewBox="0 0 ([0-9.]+) ([0-9.]+)"/.exec(source);
						if (!frame) throw new Error('The native export is missing its reference bounds.');
						return {
							...item,
							frameWidth: Number(frame[1]) / 72,
							frameHeight: Number(frame[2]) / 72,
							png: (await readFile(join(directory!, item.name + '.png'))).toString('base64'),
						};
					}),
			);
			test.skip(samples.length === 0, 'The native capture has no cases for this outline group.');
			await page.goto(framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`);
			await page.locator('#file').setInputFiles(join(directory!, 'gradient-raster.vsdx'));
			await expect(page.locator('#file-name')).toHaveText('gradient-raster.vsdx');
			const results = await page.evaluate(async (samples) => {
				const load = (path: string) => import(/* @vite-ignore */ path);
				const { renderPage, exportPageSvg } = await load('/test-api.js');
				const model = (
					document.querySelector('visio-viewer') as unknown as { document: VisioDocument }
				).document;
				const raster = async (url: string, width = 288, height = 144, x = 0, y = 0) => {
					const image = new Image();
					image.src = url;
					await image.decode();
					const canvas = document.createElement('canvas');
					canvas.width = 288;
					canvas.height = 144;
					const context = canvas.getContext('2d')!;
					context.fillStyle = 'white';
					context.fillRect(0, 0, 288, 144);
					context.drawImage(image, x, y, width, height);
					return context.getImageData(0, 0, 288, 144).data;
				};
				const results = [];
				for (const sample of samples) {
					const copy = structuredClone(model);
					copy.pages = [copy.pages[0]!];
					const view = copy.pages[0]!;
					view.shapes = view.shapes.filter((shape) => shape.id === sample.shapeId);
					const shape = view.shapes[0]!;
					const scaleX = 288 / sample.frameWidth,
						scaleY = 144 / sample.frameHeight;
					// Native bounds include the saved line-width margin even when the line is hidden.
					// SVG supplies only this export frame; PNG supplies every reference color.
					const left = shape.transform[4] - (sample.frameWidth - shape.width) / 2;
					const top =
						view.height -
						shape.transform[5] -
						shape.height -
						(sample.frameHeight - shape.height) / 2;
					const reference = await raster(`data:image/png;base64,${sample.png}`);
					// The capture script uses these exact outlines. Exclude their contour separately
					// from the fill benchmark, including diagonal edges and the reentrant notch.
					const corners = sample.outline?.length
						? sample.outline.map(([x, y]) => [x, 1 - y])
						: sample.kind === 'triangle'
							? [
									[0, 1],
									[1, 1],
									[0.5, 0],
								]
							: sample.kind === 'notched'
								? [
										[0, 1],
										[1, 1],
										[1, 0.6],
										[0.5, 0.6],
										[0.5, 0],
										[0, 0],
									]
								: [
										[0, 1],
										[1, 1],
										[1, 0],
										[0, 0],
									];
					const marginX = ((sample.frameWidth - shape.width) * scaleX) / 2,
						marginY = ((sample.frameHeight - shape.height) * scaleY) / 2;
					const vertices = corners.map(([x, y]) => [
						marginX + x! * shape.width * scaleX,
						marginY + y! * shape.height * scaleY,
					]);
					const interior = (x: number, y: number) => {
						if (sample.direction !== 13 || sample.kind === 'rectangle') return true;
						if (sample.kind === 'ellipse')
							return (
								Math.hypot(
									(x - marginX) / (shape.width * scaleX) - 0.5,
									(y - marginY) / (shape.height * scaleY) - 0.5,
								) < 0.4
							);
						let inside = false;
						for (let i = 0; i < vertices.length; i++) {
							const [ax, ay] = vertices[i]!,
								[bx, by] = vertices[(i + 1) % vertices.length]!;
							const dx = bx! - ax!,
								dy = by! - ay!;
							const t = Math.max(
								0,
								Math.min(1, ((x - ax!) * dx + (y - ay!) * dy) / (dx * dx + dy * dy)),
							);
							if (Math.hypot(x - ax! - t * dx, y - ay! - t * dy) <= 8) return false;
							if (ay! > y !== by! > y && x < ax! + ((y - ay!) * dx) / dy) inside = !inside;
						}
						return inside;
					};
					const live = renderPage(copy, view);
					try {
						for (const source of [
							new XMLSerializer().serializeToString(live.svg),
							exportPageSvg(copy, 0).svg,
						]) {
							const url = URL.createObjectURL(new Blob([source], { type: 'image/svg+xml' }));
							try {
								const actual = await raster(
									url,
									view.width * scaleX,
									view.height * scaleY,
									-left * scaleX,
									-top * scaleY,
								);
								let maximum = 0,
									total = 0,
									channels = 0;
								// Interior benchmark excludes the export bounds and outer-edge antialiasing.
								for (let y = 8; y < 136; y++)
									for (let x = 8; x < 280; x++) {
										if (!interior(x + 0.5, y + 0.5)) continue;
										for (let c = 0; c < 4; c++) {
											const i = 4 * (y * 288 + x) + c,
												delta = Math.abs(actual[i]! - reference[i]!);
											maximum = Math.max(maximum, delta);
											total += delta;
											channels++;
										}
									}
								results.push({
									name: sample.name,
									direction: sample.direction,
									alpha: sample.alpha,
									stopCount: sample.stopCount,
									maximum,
									mean: total / channels,
									pixels: channels / 4,
								});
							} finally {
								URL.revokeObjectURL(url);
							}
						}
					} finally {
						live.dispose();
					}
				}
				return results;
			}, samples);
			const output = test.info().outputPath('native-raster-differences.json');
			await writeFile(output, JSON.stringify(results, null, 2));
			await test
				.info()
				.attach('native-raster-differences', { path: output, contentType: 'application/json' });
			expect(results).toHaveLength(samples.length * 2);
			for (const item of results) expect(item.pixels, item.name).toBeGreaterThan(1000);
			// Mark only the measured fidelity gate, after successful import/render/artifact checks.
			test.fail(starOnly, 'Native star pixels exceed the existing path-fill fidelity bound.');
			// This bounds the measured improvement; nonzero differences remain parity gaps.
			for (const item of results) {
				expect(item.maximum, item.name).toBeLessThanOrEqual(item.direction === 13 ? 10 : 7);
				expect(item.mean, item.name).toBeLessThan(
					item.direction === 13 ? 2.5 : item.alpha || item.stopCount > 2 ? 1.5 : 1,
				);
			}
		});
	}
}
