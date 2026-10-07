import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { editVsdx, parseVsdx } from 'ooxml-core/visio';
import { nativeSvgLineEndpoints } from './native-line-svg';

const directory = process.env.VISIO_NATIVE_LINE_CREATION_DIR;
async function assertPaintedLine(line: import('@playwright/test').Locator) {
	await expect(line).toHaveCount(1);
	expect(
		await line.evaluate((node) => {
			const path = node.querySelector<SVGPathElement>('path')!;
			const style = getComputedStyle(path);
			return (
				style.stroke !== 'none' &&
				Number(style.opacity) > 0 &&
				style.visibility === 'visible' &&
				style.display !== 'none' &&
				path.getTotalLength() > 0
			);
		}),
	).toBe(true);
}
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: draws native straight lines, preserves formulas and saves history`, async ({
		page,
	}) => {
		test.skip(!directory, 'Set VISIO_NATIVE_LINE_CREATION_DIR to a GridAligned native capture.');
		const bytes = await readFile(join(directory!, 'moved.vsdx'));
		const native = await parseVsdx(bytes);
		const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
			cases: {
				shapeId: string;
				after: Record<string, { value: number }>;
				afterTransform: number[];
			}[];
		};
		const cleared = await editVsdx(
			bytes,
			native.pages[0]!.shapes.map((shape) => ({
				type: 'delete-shape',
				pageId: native.pages[0]!.id,
				shapeId: shape.id,
			})),
		);
		await page.goto(framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`);
		await page
			.locator('#file')
			.setInputFiles({
				name: 'blank-line.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: Buffer.from(cleared.bytes),
			});
		await expect(page.locator('#file-name')).toHaveText('blank-line.vsdx');
		const viewer = page.locator('visio-viewer');
		const viewport = viewer.locator('.viewport');
		const reference = await nativeSvgLineEndpoints(
			page,
			await readFile(join(directory!, 'moved-page.svg'), 'utf8'),
		);
		await viewport.focus();
		await viewport.press('Control+6');
		await expect(viewport).toHaveAttribute('data-tool', 'line');
		for (const item of evidence.cases) {
			const points = await viewer.evaluate((node, item) => {
				const element = node as unknown as {
					document: import('ooxml-core/visio').VisioDocument;
					shadowRoot: ShadowRoot;
				};
				const model = element.document.pages[0]!;
				const matrix = element.shadowRoot
					.querySelector<SVGSVGElement>('svg.paper')!
					.getScreenCTM()!;
				return ['Begin', 'End'].map((prefix) => {
					const point = new DOMPoint(
						item.after[prefix + 'X']!.value,
						model.height - item.after[prefix + 'Y']!.value,
					).matrixTransform(matrix);
					return { x: point.x, y: point.y };
				});
			}, item);
			await page.mouse.move(points[0]!.x, points[0]!.y);
			await page.mouse.down();
			await page.mouse.move(points[1]!.x, points[1]!.y, { steps: 4 });
			await expect(viewer.locator('line.draw-preview')).toHaveCount(1);
			expect(
				await viewer.locator('line.draw-preview').evaluate((node) => {
					const style = getComputedStyle(node);
					return (
						style.stroke !== 'none' &&
						style.visibility === 'visible' &&
						style.display !== 'none' &&
						(node as SVGLineElement).getTotalLength() > 0
					);
				}),
			).toBe(true);
			await page.mouse.up();
			await expect(viewer.locator('.draw-preview')).toHaveCount(0);
			const line = viewer.locator(`[data-shape-id="${item.shapeId}"]`);
			await assertPaintedLine(line);
			const pose = (await line.getAttribute('transform'))!
				.slice(7, -1)
				.trim()
				.split(/[\s,]+/)
				.map(Number);
			for (let i = 0; i < 6; i++) expect(pose[i]).toBeCloseTo(item.afterTransform[i]!, 12);
			await viewport.focus();
			await viewport.press('Control+z');
			await expect(line).toHaveCount(0);
			await viewport.press('Control+y');
			await assertPaintedLine(line);
		}
		const saved = Buffer.from(
			await viewer.evaluate((element) =>
				Array.from(
					(element as unknown as { exportVsdx(): { bytes: Uint8Array } }).exportVsdx().bytes,
				),
			),
		);
		const model = await parseVsdx(saved);
		for (const shape of model.pages[0]!.shapes) {
			const expected = native.pages[0]!.shapes.find((item) => item.id === shape.id)!;
			expect(shape.geometry).toEqual(expected.geometry);
			expect(shape.style).toEqual(expected.style);
			const nativeLine = reference.lines.find((line) => line.id === shape.id)!;
			const [a, b, , , x, y] = shape.transform;
			expect(x).toBeCloseTo(nativeLine.begin.x, 3);
			expect(y).toBeCloseTo(nativeLine.begin.y, 3);
			expect(x + a * shape.width).toBeCloseTo(nativeLine.end.x, 3);
			expect(y + b * shape.width).toBeCloseTo(nativeLine.end.y, 3);
		}
		await page
			.locator('#file')
			.setInputFiles({
				name: 'created-lines.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: saved,
			});
		await expect(page.locator('#file-name')).toHaveText('created-lines.vsdx');
		await expect(viewer.locator('[data-shape-id]')).toHaveCount(4);
		await viewport.focus();
		await viewport.press('Control+6');
		const cancel = await viewer.evaluate((element) => {
			const svg = (element as HTMLElement).shadowRoot!.querySelector<SVGSVGElement>('svg.paper')!;
			const point = new DOMPoint(6, 4).matrixTransform(svg.getScreenCTM()!);
			return { x: point.x, y: point.y };
		});
		await page.mouse.move(cancel.x, cancel.y);
		await page.mouse.down();
		await page.mouse.move(cancel.x + 25, cancel.y);
		await expect(viewer.locator('line.draw-preview')).toHaveCount(1);
		await viewport.press('Escape');
		await page.mouse.up();
		await expect(viewer.locator('.draw-preview')).toHaveCount(0);
		expect(
			Buffer.from(
				await viewer.evaluate((element) =>
					Array.from(
						(element as unknown as { exportVsdx(): { bytes: Uint8Array } }).exportVsdx().bytes,
					),
				),
			),
		).toEqual(saved);
	});
}
