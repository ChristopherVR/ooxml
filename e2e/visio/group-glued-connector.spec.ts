import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { openDemo } from './demo-page';
import { handlePoint, inventory } from './selection-resize-fixture';

/**
 * A group of two rectangles with a Visio Dynamic connector glued to one member (saved by Visio 16,
 * see src/core/visio/edit-group-glue-native.test.ts). Moving or resizing the group reroutes the
 * connector onto the member's new place.
 */
const fixtureBytes = readFileSync(
	new URL('../../src/core/visio/__fixtures__/group-glued-connector.vsdx', import.meta.url),
);

async function connectorEnds(bytes: number[]): Promise<number[]> {
	const zip = await JSZip.loadAsync(Buffer.from(bytes));
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	const shape = /<Shape ID=['"]5['"][\s\S]*?<\/Shape>/.exec(xml)![0];
	return ['BeginX', 'BeginY', 'EndX', 'EndY'].map((name) =>
		Number(new RegExp(`<Cell N=['"]${name}['"] V=['"]([^'"]*)['"]`).exec(shape)![1]),
	);
}

for (const framework of ['vanilla', 'react']) {
	test(`${framework}: a connector glued to a group member follows the group`, async ({ page }) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		await page.setViewportSize({ width: 1600, height: 1000 });
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		const viewer = page.locator('visio-viewer');
		await page.locator('#file').setInputFiles({
			name: 'group-glued.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: fixtureBytes,
		});
		await expect.poll(async () => (await inventory(viewer)).shapes.length).toBe(3);
		const start = await inventory(viewer);
		const begin = await connectorEnds(start.bytes);
		expect(begin[0]).toBeCloseTo(2.5, 6);
		expect(begin[1]).toBeCloseTo(4, 6);
		// Select the group through the keyboard, then nudge it.
		const group = viewer.locator('svg.paper [data-shape-id="4"]');
		await group.focus();
		await group.press('Enter');
		await expect(viewer.locator('[data-resize-handle]')).toHaveCount(8);
		await viewer.locator('.viewport').press('ArrowRight');
		await expect
			.poll(async () => (await connectorEnds((await inventory(viewer)).bytes))[0])
			.toBeGreaterThan(2.5);
		const moved = await inventory(viewer);
		expect(moved.error).toBe('');
		const nudged = await connectorEnds(moved.bytes);
		const groupMove = moved.shapes.find((shape) => shape.children.length)!.transform[4]!;
		const groupBefore = start.shapes.find((shape) => shape.children.length)!.transform[4]!;
		expect(nudged[0]! - begin[0]!).toBeCloseTo(groupMove - groupBefore, 6);
		expect(nudged[1]).toBeCloseTo(begin[1]!, 6);
		// Drag the corner handle: the connector lands on the resized member's top edge.
		const handle = await handlePoint(viewer);
		await page.mouse.move(handle.x, handle.y);
		await page.mouse.down();
		await page.mouse.move(handle.x + handle.dx, handle.y + handle.dy, { steps: 4 });
		await page.mouse.up();
		await expect
			.poll(async () => (await inventory(viewer)).shapes.find((s) => s.children.length)!.width)
			.toBeGreaterThan(3.5);
		const resized = await inventory(viewer);
		expect(resized.error).toBe('');
		const outer = resized.shapes.find((shape) => shape.children.length)!;
		const member = outer.children.find((item) => item.id === '1')!;
		// The member's top-edge midpoint through the group's transform, then the connector end.
		const [ga, gb, gc, gd, ge, gf] = outer.transform as [
			number,
			number,
			number,
			number,
			number,
			number,
		];
		const [ma, mb, mc, md, me, mf] = member.transform as [
			number,
			number,
			number,
			number,
			number,
			number,
		];
		const local = {
			x: ma * (member.width / 2) + mc * member.height + me,
			y: mb * (member.width / 2) + md * member.height + mf,
		};
		const world = { x: ga * local.x + gc * local.y + ge, y: gb * local.x + gd * local.y + gf };
		const ends = await connectorEnds(resized.bytes);
		expect(ends[0]).toBeCloseTo(world.x, 5);
		expect(ends[1]).toBeCloseTo(world.y, 5);
		expect(errors).toEqual([]);
	});
}
