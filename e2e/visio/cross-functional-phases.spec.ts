import { readFileSync } from 'node:fs';
import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

const TAB = '[data-ribbon-tab="cross-functional"]';
/** The chart Visio 16 builds from its Cross-Functional Flowchart template (frame 1, lists 4 and 5). */
const native = readFileSync(
	new URL('../../src/core/visio/__fixtures__/cross-functional-native.vsdx', import.meta.url),
);

/** The chart as the viewer's model holds it: options, phases in order, the lanes' top edge. */
function chart(viewer: Locator, frameId: string, phaseListId: string, swimListId = '4') {
	return viewer.evaluate(
		(node, ids) => {
			const shapes = (node as VisioViewerElement).controller.state.document!.pages[0]!.shapes;
			const byId = (id: string) => shapes.find((shape) => shape.id === id)!;
			const frame = byId(ids.frameId);
			const swimlane = frame.structure?.type === 'container' ? frame.structure.swimlane : undefined;
			const options = swimlane?.role === 'frame' ? swimlane.options : undefined;
			const list = byId(ids.phaseListId).structure;
			const phases = list?.type === 'list' ? list.memberIds : [];
			const top = (id: string) => byId(id).transform[5] + byId(id).height;
			const lanes = shapes.filter(
				(shape) =>
					shape.structure?.type === 'container' && shape.structure.swimlane?.role === 'lane',
			);
			const swim = byId(ids.swimListId).structure;
			return {
				options,
				rotateLabels: swim?.type === 'list' ? swim.rotateLabels : undefined,
				phases,
				phaseSpans: phases.map((id) => [byId(id).transform[4], byId(id).width]),
				laneTop: Math.max(...lanes.map((lane) => top(lane.id))),
				frameHeight: frame.height,
			};
		},
		{ frameId, phaseListId, swimListId },
	);
}

async function savedPage(viewer: Locator): Promise<string> {
	const bytes = await viewer.evaluate((node) =>
		Array.from((node as VisioViewerElement).controller.exportVsdx().bytes),
	);
	return (await JSZip.loadAsync(Uint8Array.from(bytes)))
		.file('visio/pages/page1.xml')!
		.async('string');
}

async function openChartTab(viewer: Locator, laneId: string): Promise<Locator> {
	await viewer
		.locator(`svg.paper [data-shape-id="${laneId}"]`)
		.first()
		.click({ force: true, position: { x: 10, y: 10 } });
	await viewer.locator('office-ui-ribbon #tab-cross-functional').click();
	return viewer.locator(TAB);
}

test('Insert Separator and the chart options work on a chart Visio built', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Visio chart.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: native,
	});
	await expect(page.locator('#file-name')).toHaveText('Visio chart.vsdx');
	const before = await chart(viewer, '1', '5');
	expect(before.phases).toEqual(['12']);
	expect(before.options).toMatchObject({ showTitle: true, showSeparators: true, style: 1 });

	const panel = await openChartTab(viewer, '6');
	// Separator splits the template's one phase: the new phase comes first, edge to edge.
	await panel.getByRole('button', { name: 'Separator', exact: true }).click();
	await expect.poll(async () => (await chart(viewer, '1', '5')).phases.length).toBe(2);
	const split = await chart(viewer, '1', '5');
	expect(split.phases[1]).toBe('12');
	const [first, second] = split.phaseSpans as [number, number][];
	expect(first![0]! + first![1]!).toBeCloseTo(second![0]!);
	expect(first![1]! + second![1]!).toBeCloseTo(before.phaseSpans[0]![1]!);
	let xml = await savedPage(viewer);
	expect(xml).toMatch(/DEPENDSON\(2,Sheet\.\d+!SheetRef\(\),Sheet\.12!SheetRef\(\)\)/);
	// One undo step puts the phase back.
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
	await expect.poll(async () => (await chart(viewer, '1', '5')).phases).toEqual(['12']);
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.redo());
	await expect.poll(async () => (await chart(viewer, '1', '5')).phases.length).toBe(2);

	// Show Title Bar: the frame loses the bar, the lanes stay.
	await panel.locator('[data-check="swimlane-title"]').click();
	await expect.poll(async () => (await chart(viewer, '1', '5')).options?.showTitle).toBe(false);
	const noTitle = await chart(viewer, '1', '5');
	expect(noTitle.frameHeight).toBeLessThan(before.frameHeight - 0.3);
	expect(noTitle.laneTop).toBeCloseTo(before.laneTop);

	// Show Separator Bar off: the lanes rise by the band, the frame shrinks.
	await panel.locator('[data-check="swimlane-separators"]').click();
	await expect
		.poll(async () => (await chart(viewer, '1', '5')).options?.showSeparators)
		.toBe(false);
	const noPhases = await chart(viewer, '1', '5');
	expect(noPhases.laneTop).toBeGreaterThan(before.laneTop + 0.15);
	expect(noPhases.frameHeight).toBeLessThan(noTitle.frameHeight - 0.15);
	await panel.locator('[data-check="swimlane-separators"]').click();
	await panel.locator('[data-check="swimlane-title"]').click();
	await expect
		.poll(async () => (await chart(viewer, '1', '5')).frameHeight)
		.toBeGreaterThan(before.frameHeight - 0.01);
	const restored = await chart(viewer, '1', '5');
	expect(restored.laneTop).toBeCloseTo(before.laneTop);
	expect(restored.frameHeight).toBeCloseTo(before.frameHeight);

	// Style 3 and Rotate Lane Labels reach the saved cells Visio's masters read.
	await panel.getByRole('button', { name: 'Style', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Style 3"]`).click();
	await expect.poll(async () => (await chart(viewer, '1', '5')).options?.style).toBe(3);
	await panel.getByRole('button', { name: 'Rotate Lane Labels', exact: true }).click();
	await expect.poll(async () => (await chart(viewer, '1', '5')).rotateLabels).toBe(true);
	xml = await savedPage(viewer);
	expect(xml).toContain('/stg1=3');
	expect(xml).toContain('/stg6=1');

	// Direction cannot be changed on Visio's chart: the chart is left as it is.
	await panel.getByRole('button', { name: 'Direction', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Vertical"]`).click();
	await expect(viewer.getByText(/Direction of a chart Visio built/).first()).toBeVisible();
	expect((await chart(viewer, '1', '5')).options?.direction).toBe('horizontal');
	expect(errors).toEqual([]);
});

test('Insert Separator and the chart options work on a chart the editor creates', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer.locator('[data-backstage-action="new-template:cross-functional-flowchart"]').click();
	await expect(viewer.locator('svg.paper')).toContainText('Place order');
	const panel = await openChartTab(viewer, '11');
	const before = await chart(viewer, '50', '52');

	await panel.getByRole('button', { name: 'Separator', exact: true }).click();
	await expect
		.poll(async () => (await chart(viewer, '50', '52')).phases.length)
		.toBe(before.phases.length + 1);
	await panel.locator('[data-check="swimlane-separators"]').click();
	await expect
		.poll(async () => (await chart(viewer, '50', '52')).options?.showSeparators)
		.toBe(false);
	await panel.locator('[data-check="swimlane-title"]').click();
	await expect.poll(async () => (await chart(viewer, '50', '52')).options?.showTitle).toBe(false);
	await panel.getByRole('button', { name: 'Style', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Style 4"]`).click();
	await expect.poll(async () => (await chart(viewer, '50', '52')).options?.style).toBe(4);
	const xml = await savedPage(viewer);
	for (const row of ['visShowPhase', 'visShowTitle'])
		expect(xml).toMatch(new RegExp(`N="${row}"><Cell N="Value" V="0"(?! U=)`));
	// Direction still works on a chart the editor created.
	await panel.getByRole('button', { name: 'Direction', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Vertical"]`).click();
	await expect
		.poll(async () => (await chart(viewer, '50', '52')).options?.direction)
		.toBe('vertical');
	expect(errors).toEqual([]);
});
