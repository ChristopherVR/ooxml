import { test, expect, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

async function client(viewer: Locator, x: number, y: number) {
	const paper = (await viewer.locator('svg.paper').boundingBox())!;
	return { x: paper.x + (x / 8.5) * paper.width, y: paper.y + ((11 - y) / 11) * paper.height };
}
/** How many straight runs a saved connector has: a level connector has one. */
async function runs(viewer: Locator, id: string): Promise<number> {
	const saved = await viewer.evaluate((node) => [
		...(node as VisioViewerElement).controller.exportVsdx().bytes,
	]);
	const zip = await JSZip.loadAsync(Buffer.from(saved));
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	const shape = new RegExp(`<Shape ID="${id}"[\\s\\S]*?</Shape>`, 'u').exec(xml)![0];
	return (shape.match(/T="LineTo"/gu) ?? []).length;
}

test('a shape dragged onto a right-angle connector pushes the connector aside', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1500, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await viewer.evaluate(async (node) => {
		const { controller } = node as VisioViewerElement;
		await controller.createBlankDrawing();
		const pageId = controller.state.document!.pages[0]!.id;
		const box = (shapeId: string, x: number, y: number) =>
			({ type: 'create-rectangle', pageId, shapeId, x, y, width: 1, height: 1 }) as const;
		const connect = (shapeId: string, begin: string, end: string) =>
			({
				type: 'create-line',
				pageId,
				shapeId,
				beginX: 0,
				beginY: 0,
				endX: 1,
				endY: 1,
				connect: { begin, end },
				route: 'right-angle',
			}) as const;
		// Connecting 4 to 5 makes both placeable, as in Visio; 3 joins 1 and 2 along a level run.
		await controller.applyEdits([
			box('1', 1.5, 4),
			box('2', 7, 4),
			box('4', 4.25, 8),
			box('5', 7, 8),
		]);
		await controller.applyEdits([connect('6', '4', '5')]);
		await controller.applyEdits([connect('3', '1', '2')]);
	});
	expect(await runs(viewer, '3')).toBe(1);
	const from = await client(viewer, 4.25, 8);
	const to = await client(viewer, 4.25, 4);
	await page.mouse.move(from.x, from.y);
	await page.mouse.down();
	await page.mouse.move(to.x, to.y, { steps: 8 });
	await page.mouse.up();
	await expect
		.poll(() => viewer.evaluate((node) => (node as VisioViewerElement).controller.state.edit))
		.toMatchObject({ busy: false });
	// The connector went around the shape; 6 followed it, both in the one move.
	await expect.poll(() => runs(viewer, '3')).toBeGreaterThan(1);
	await expect.poll(() => runs(viewer, '6')).toBeGreaterThan(0);
	// One undo puts the shape and the connector back.
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
	await expect.poll(() => runs(viewer, '3')).toBe(1);
	expect(errors).toEqual([]);
});
