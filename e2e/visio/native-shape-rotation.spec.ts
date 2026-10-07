import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { nativeSvgLineEndpoints } from './native-line-svg';

for (const variable of [
	'VISIO_NATIVE_ROTATE_DIR',
	'VISIO_NATIVE_ROTATE_DOUBLE_DIR',
	'VISIO_NATIVE_ROTATE_HALF_DIR',
	'VISIO_NATIVE_ROTATE_TRIPLE_DIR',
]) {
	const directory = process.env[variable];
	for (const kind of ['rectangle', 'ellipse'])
		for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
			test(`${framework}: rotates native ${kind}, preserves pin and saves history (${variable})`, async ({
				page,
			}) => {
				test.skip(!directory, `Set ${variable} to a native rotated shape capture.`);
				const source = await readFile(join(directory!, 'ellipse-edited.vsdx'));
				const native = await parseVsdx(await readFile(join(directory!, 'rotated.vsdx')));
				const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
					rotated: Record<
						string,
						{ shapeId: string; cells: Record<string, { value: number }>; transform: number[] }
					>;
				};
				const reference = evidence.rotated[kind]!;
				const expected = native.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
				const ratio = native.pages[0]!.drawingToPageScale ?? 1;
				await page.goto(
					framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
				);
				await page.locator('#file').setInputFiles({
					name: 'native-rotation.vsdx',
					mimeType: 'application/vnd.ms-visio.drawing',
					buffer: source,
				});
				await expect(page.locator('#file-name')).toHaveText('native-rotation.vsdx');
				const viewer = page.locator('visio-viewer'),
					viewport = viewer.locator('.viewport');
				const shape = viewer.locator(`[data-shape-id="${reference.shapeId}"]`);
				await expect(shape).toBeVisible();
				const before = (await shape.getAttribute('transform'))!;
				await shape.focus();
				await shape.press('Enter');
				const controls = viewer.locator('.edit-controls');
				await controls.locator('summary').click();
				const rotate = controls.getByRole('button', { name: 'Rotate selected', exact: true });
				await expect(rotate).toBeDisabled();
				await controls
					.getByLabel('Angle (degrees)', { exact: true })
					.fill(String((reference.cells.Angle!.value * 180) / Math.PI));
				await expect(rotate).toBeEnabled();
				await rotate.click();
				await expect(shape).not.toHaveAttribute('transform', before);
				const after = (await shape.getAttribute('transform'))!;
				await viewport.focus();
				await viewport.press('Control+z');
				await expect(shape).toHaveAttribute('transform', before);
				await viewport.press('Control+y');
				await expect(shape).toHaveAttribute('transform', after);
				const saved = Buffer.from(
					await viewer.evaluate((element) =>
						Array.from(
							(element as unknown as { exportVsdx(): { bytes: Uint8Array } }).exportVsdx().bytes,
						),
					),
				);
				const model = await parseVsdx(saved);
				const actual = model.pages[0]!.shapes.find((shape) => shape.id === reference.shapeId)!;
				expect(actual.geometry).toEqual(expected.geometry);
				expect(actual.style).toEqual(expected.style);
				expect(model.pages[0]!.drawingToPageScale ?? 1).toBe(ratio);
				for (let i = 0; i < 6; i++)
					expect(actual.transform[i]).toBeCloseTo(
						reference.transform[i]! * (i >= 4 ? ratio : 1),
						12,
					);
				await page.locator('#file').setInputFiles({
					name: 'rotated-copy.vsdx',
					mimeType: 'application/vnd.ms-visio.drawing',
					buffer: saved,
				});
				await expect(page.locator('#file-name')).toHaveText('rotated-copy.vsdx');
				await expect(shape).toHaveAttribute('transform', after);
				const svg = await nativeSvgLineEndpoints(
					page,
					await readFile(join(directory!, 'rotated-page.svg'), 'utf8'),
				);
				const nativeShape = svg.lines.find((shape) => shape.id === reference.shapeId)!;
				const rendered = await shape
					.locator('path')
					.first()
					.evaluate((path) => {
						const primitive = path as SVGGraphicsElement;
						const paper = primitive.closest('svg') as SVGSVGElement;
						const matrix = paper.getScreenCTM()!.inverse().multiply(primitive.getScreenCTM()!);
						const box = primitive.getBBox();
						const points = [
							[box.x, box.y],
							[box.x + box.width, box.y],
							[box.x, box.y + box.height],
							[box.x + box.width, box.y + box.height],
						].map(([x, y]) => new DOMPoint(x, y).matrixTransform(matrix));
						const style = getComputedStyle(path);
						return {
							width: Math.max(...points.map((p) => p.x)) - Math.min(...points.map((p) => p.x)),
							height: Math.max(...points.map((p) => p.y)) - Math.min(...points.map((p) => p.y)),
							paint: {
								fill: style.fill,
								stroke: style.stroke,
								fillOpacity: style.fillOpacity,
								strokeOpacity: style.strokeOpacity,
							},
						};
					});
				expect(rendered.width).toBeCloseTo(nativeShape.bounds.width, 3);
				expect(rendered.height).toBeCloseTo(nativeShape.bounds.height, 3);
				expect(rendered.paint).toEqual(nativeShape.paint);
			});
		}
}
