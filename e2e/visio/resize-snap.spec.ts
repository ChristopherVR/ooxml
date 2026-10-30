import { expect, test, type Locator } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';
import { fixture, inventory } from './selection-resize-fixture';

const waitForIdle = (viewer: Locator) =>
	expect
		.poll(() =>
			viewer.evaluate((node) => {
				const state = (node as VisioViewerElement).controller.state;
				return state.loading || state.edit.busy;
			}),
		)
		.toBe(false);

const handlePosition = (viewer: Locator, id: string) =>
	viewer.locator(`[data-resize-handle="${id}"]`).evaluate((node) => {
		const circle = node as SVGCircleElement,
			point = new DOMPoint(circle.cx.baseVal.value, circle.cy.baseVal.value).matrixTransform(
				circle.getScreenCTM()!,
			);
		return { x: point.x, y: point.y };
	});

/** Client pixels for a page-inch vector, for a page with no rotation. */
const clientDelta = (viewer: Locator, x: number) =>
	viewer.locator('svg.paper').evaluate((node, dx) => {
		const ctm = (node as SVGSVGElement).getScreenCTM()!;
		return { x: ctm.a * dx, y: ctm.b * dx };
	}, x);

test('vanilla: a resize edge snaps to a neighbour, shows the alignment hint and commits the snap', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'resize-snap.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await fixture(),
	});
	await expect(page.locator('#file-name')).toHaveText('resize-snap.vsdx');
	await waitForIdle(viewer);
	await viewer.locator('svg.paper [data-shape-id="1"]').click();
	await expect(viewer.locator('[data-resize-handle]')).toHaveCount(8);
	await viewer.getByRole('tab', { name: 'Home', exact: true }).click();
	const before = await inventory(viewer);
	const shapeOne = before.shapes[0]!;
	const rightEdge = shapeOne.pin!.pinX + shapeOne.width / 2;
	const neighbourLeft = before.shapes[1]!.pin!.pinX - before.shapes[1]!.width / 2;
	// Drag the east handle to 0.04 inch short of the neighbour's left edge, inside the snap threshold.
	const travel = neighbourLeft - 0.04 - rightEdge;
	const start = await handlePosition(viewer, 'e');
	const delta = await clientDelta(viewer, travel);
	await page.mouse.move(start.x, start.y);
	await page.mouse.down();
	await page.mouse.move(start.x + delta.x, start.y + delta.y, { steps: 6 });
	// The alignment line is drawn while the edge is held on the neighbour's edge.
	await expect(viewer.locator('[data-snap-hints] line')).toHaveCount(1);
	await page.mouse.up();
	await expect
		.poll(async () => (await inventory(viewer)).shapes[0]!.width)
		.toBeGreaterThan(shapeOne.width);
	const resized = await inventory(viewer);
	// The right edge lands on the neighbour's left edge; the left edge stays where it was.
	const newRight = resized.shapes[0]!.pin!.pinX + resized.shapes[0]!.width / 2;
	expect(newRight).toBeCloseTo(neighbourLeft, 6);
	await expect(viewer.locator('[data-snap-hints]')).toHaveCount(0);
	expect(errors).toEqual([]);
});
