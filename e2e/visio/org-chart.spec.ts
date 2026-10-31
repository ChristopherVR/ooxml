import { expect, test, type Locator, type Page } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';
import { taskPane } from './ribbon';

const TAB = '[data-ribbon-tab="org-chart"]';
/** Shape IDs of the template's positions: a Belt master instance takes sixteen IDs. */
const MANAGER = '17';
const LEAF = '65';

interface Chart {
	positions: { id: string; name: string; text: string; x: number; y: number }[];
	connectors: number;
}
/** The shapes that carry the org chart Name row, with the master they are an instance of. */
async function chart(viewer: Locator): Promise<Chart> {
	return viewer.evaluate((node) => {
		const page = (node as VisioViewerElement).controller.state.document!.pages[0]!;
		return {
			positions: page.shapes
				.filter((shape) => shape.shapeData?.some((item) => item.name === 'Name'))
				.map((shape) => ({
					id: shape.id,
					name: shape.name,
					text: shape.text.plainText,
					x: shape.transform[4] + shape.width / 2,
					y: shape.transform[5] + shape.height / 2,
				})),
			connectors: page.connectors.length / 2,
		};
	});
}

async function openTemplate(page: Page): Promise<Locator> {
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer.locator('[data-backstage-action="new-template:organization-chart"]').click();
	await expect(viewer.locator('svg.paper')).toContainText('Chief Executive');
	return viewer;
}
const pick = (viewer: Locator, id: string) =>
	viewer
		.locator(`svg.paper [data-shape-id="${id}"]`)
		.click({ force: true, position: { x: 10, y: 10 } });
const undo = (viewer: Locator) =>
	viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());

test('the New tiles show thumbnails and the org chart gets its own tab', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();

	// Every template tile shows an image rendered from its template.
	const tiles = viewer.locator('.backstage-templates .backstage-action');
	await expect(tiles).toHaveCount(5);
	await expect(viewer.locator('.backstage-templates img.backstage-thumb')).toHaveCount(5);
	for (const image of await viewer.locator('.backstage-templates img.backstage-thumb').all())
		await expect
			.poll(() => image.evaluate((node) => (node as HTMLImageElement).naturalWidth))
			.toBeGreaterThan(0);

	await viewer.locator('[data-backstage-action="new-template:organization-chart"]').click();
	await expect(viewer.locator('svg.paper')).toContainText('Chief Executive');
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(1);

	// The positions are instances of Visio's Executive, Manager and Position Belt masters, which
	// the Document Stencil lists with the Dynamic connector.
	const before = await chart(viewer);
	expect(before.positions.map((item) => item.name.replace(/\.\d+$/, ''))).toEqual([
		'Executive Belt',
		'Manager Belt',
		'Manager Belt',
		'Manager Belt',
		'Position Belt',
		'Position Belt',
	]);
	await expect(viewer.getByText('Executive Belt').first()).toBeVisible();

	// Selecting a position shows the Org Chart tab.
	await pick(viewer, MANAGER);
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(0);
	await viewer.locator('office-ui-ribbon #tab-org-chart').click();
	const panel = viewer.locator(TAB);
	expect((await chart(viewer)).positions).toHaveLength(6);

	await panel.locator('[command="org-subordinate"]').click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);
	expect((await chart(viewer)).connectors).toBe(6);
	// One undo step per command.
	await undo(viewer);
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.redo());
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);

	// Delete the new position: it and its connector go.
	const added = (await chart(viewer)).positions.at(-1)!.id;
	await pick(viewer, added);
	await panel.locator('[command="org-delete"]').click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	expect((await chart(viewer)).connectors).toBe(5);
	expect(errors).toEqual([]);
});

test('an assistant sits beside its manager and the Delete key deletes the position in one undo step', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	const viewer = await openTemplate(page);
	await pick(viewer, MANAGER);
	await viewer.locator('office-ui-ribbon #tab-org-chart').click();
	await viewer.locator(`${TAB} [command="org-assistant"]`).click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);
	const state = await chart(viewer);
	const manager = state.positions.find((item) => item.id === MANAGER)!;
	const assistant = state.positions.at(-1)!;
	expect(assistant.name).toMatch(/^Assistant Belt/);
	// Beside the manager's line, not under the manager, and above the manager's subordinates.
	expect(assistant.x).toBeGreaterThan(manager.x + 0.8);
	expect(assistant.y).toBeLessThan(manager.y);
	const leaf = state.positions.find((item) => item.id === LEAF)!;
	expect(assistant.y).toBeGreaterThan(leaf.y);

	// Delete on a position is Delete Position: its subordinates join its manager.
	await pick(viewer, MANAGER);
	await viewer.locator('svg.paper').focus();
	await page.keyboard.press('Delete');
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	const after = await viewer.evaluate((node) => {
		const doc = (node as VisioViewerElement).controller.state.document!.pages[0]!;
		return doc.connectors
			.filter((row) => row.fromCell === 'EndX')
			.map((row) => row.toShapeId)
			.sort();
	});
	// The assistant and both leaves now report to the Chief Executive: three managers' worth of
	// lines minus the deleted one, plus three joins.
	expect(after).toContain(LEAF);
	expect(after).toContain(assistant.id);
	expect((await chart(viewer)).positions.some((item) => item.id === MANAGER)).toBe(false);
	// One undo brings the position, its connectors and its assistant's line back.
	await undo(viewer);
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);
	expect((await chart(viewer)).connectors).toBe(6);
	expect(errors).toEqual([]);
});

test('a Position Belt that gains a subordinate becomes a Manager Belt glued at Top and Bottom', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	const viewer = await openTemplate(page);
	// Every reporting line joins the manager's Bottom point (row 4) to the subordinate's Top (row 1).
	const glue = () =>
		viewer.evaluate((node) =>
			(node as VisioViewerElement).controller.state
				.document!.pages[0]!.connectors.filter((row) => /^(Begin|End)X$/.test(row.fromCell))
				.map((row) => `${row.fromCell}:${row.toCell}`),
		);
	expect(new Set(await glue())).toEqual(new Set(['BeginX:Connections.X4', 'EndX:Connections.X1']));

	await pick(viewer, LEAF);
	await viewer.locator('office-ui-ribbon #tab-org-chart').click();
	await viewer.locator(`${TAB} [command="org-subordinate"]`).click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);
	const state = await chart(viewer);
	// The leaf was replaced by a Manager Belt (same text, new id) and keeps its place and line.
	expect(state.positions.some((item) => item.id === LEAF)).toBe(false);
	const kinds = state.positions.map((item) => item.name.replace(/\.\d+$/, ''));
	expect(kinds.filter((name) => name === 'Manager Belt')).toHaveLength(4);
	expect(kinds.filter((name) => name === 'Position Belt')).toHaveLength(2);
	const promoted = state.positions.find(
		(item) => item.name.startsWith('Manager Belt') && !['17', '33', '49'].includes(item.id),
	)!;
	expect(promoted.text).toBe('Name\nPosition');
	expect(state.connectors).toBe(6);
	expect(new Set(await glue())).toEqual(new Set(['BeginX:Connections.X4', 'EndX:Connections.X1']));

	// Moving the promoted manager keeps its lines on the same connection points.
	await viewer.evaluate(
		(node, id) =>
			(node as VisioViewerElement).controller.applyEdits([
				{ type: 'move-shape', pageId: '0', shapeId: id, x: 9, y: 2.5 },
			]),
		promoted.id,
	);
	expect(new Set(await glue())).toEqual(new Set(['BeginX:Connections.X4', 'EndX:Connections.X1']));

	// One undo step per command: both the move and the type change step back.
	await undo(viewer);
	await undo(viewer);
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	expect((await chart(viewer)).positions.some((item) => item.id === LEAF)).toBe(true);
	expect(errors).toEqual([]);
});

test('the text of a position follows its Shape Data', async ({ page }) => {
	const viewer = await openTemplate(page);
	await pick(viewer, LEAF);
	await taskPane(viewer, 'Shape Data');
	await viewer.locator('[data-define-shape-data]').click();
	const dialog = viewer.locator('.shape-data-dialog');
	await dialog.locator('[name="row"]').selectOption('Name');
	await dialog.locator('[name="value"]').fill('Ada Lovelace');
	await dialog.locator('[label="OK"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Saved Shape Data Name/);
	await expect(viewer.locator(`svg.paper [data-shape-id="${LEAF}"]`)).toContainText('Ada Lovelace');
	expect((await chart(viewer)).positions.find((item) => item.id === LEAF)!.text).toBe(
		'Ada Lovelace\nPosition',
	);
	// A change to a row that is not shown leaves the text alone.
	await viewer.locator('[data-define-shape-data]').click();
	await dialog.locator('[name="row"]').selectOption('Telephone');
	await dialog.locator('[name="value"]').fill('555 0100');
	await dialog.locator('[label="OK"] button').click();
	await expect(viewer.locator('[data-status]')).toHaveText(/Saved Shape Data Telephone/);
	expect((await chart(viewer)).positions.find((item) => item.id === LEAF)!.text).toBe(
		'Ada Lovelace\nPosition',
	);
});

test('Import Org Chart builds a chart from a sheet with Name, Title and Reports To', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => {
		const csv =
			'Name,Title,Reports To,Department\n' +
			'Ada Lovelace,Chief Executive,,Board\n' +
			'Bo Chen,CTO,Ada Lovelace,Technology\n' +
			'Cy Okoye,CFO,Ada Lovelace,Finance\n' +
			'Di Park,Engineer,Bo Chen,Technology\n' +
			'Ed Silva,Executive Assistant,Ada Lovelace,Board\n';
		(window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => [
			{ getFile: async () => new File([csv], 'people.csv', { type: 'text/csv' }) },
		];
	});
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer.locator('[data-backstage-action="new-blank"]').click();
	await expect(viewer.locator('svg.paper')).toBeVisible();
	await viewer.getByRole('tab', { name: 'Data', exact: true }).click();
	await viewer
		.locator('[data-ribbon-tab]:not([data-contextual]) [command="org-import-data"] button')
		.click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(5);
	const result = await chart(viewer);
	expect(result.connectors).toBe(4);
	const names = result.positions.map((item) => item.name.replace(/\.\d+$/, ''));
	expect(names.sort()).toEqual([
		'Executive Belt',
		'Manager Belt',
		'Position Belt',
		'Position Belt',
		'Position Belt',
	]);
	expect(result.positions.map((item) => item.text)).toContain('Bo Chen\nCTO');
	await expect(viewer.locator('svg.paper')).toContainText('Ada Lovelace');
	// The import is one undo step.
	await undo(viewer);
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(0);
	expect(errors).toEqual([]);
});
