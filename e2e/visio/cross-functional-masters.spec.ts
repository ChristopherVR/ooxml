import { expect, test, type Locator } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

const TAB = '[data-ribbon-tab="cross-functional"]';

/** The saved drawing's masters and, per page shape, the master it is an instance of. */
async function saved(viewer: Locator) {
	const bytes = await viewer.evaluate((node) =>
		Array.from((node as VisioViewerElement).controller.exportVsdx().bytes),
	);
	const zip = await JSZip.loadAsync(Uint8Array.from(bytes));
	const masters = await zip.file('visio/masters/masters.xml')!.async('string');
	const page = await zip.file('visio/pages/page1.xml')!.async('string');
	const names = new Map<string, string>();
	for (const match of masters.matchAll(/<Master ID="(\d+)" NameU="([^"]*)"[^>]*Hidden="(\d)"/g))
		names.set(match[1]!, `${match[2]}${match[3] === '1' ? ' (hidden)' : ''}`);
	const instances = new Map<string, string>();
	for (const match of page.matchAll(/<Shape ID="(\d+)"[^>]*? Master="(\d+)"/g))
		instances.set(match[1]!, names.get(match[2]!) ?? '?');
	const rels = await zip.file('visio/pages/_rels/page1.xml.rels')!.async('string');
	return { names: [...names.values()], instances, rels };
}

test('a cross-functional flowchart the editor creates is built from the hidden flowchart masters', async ({
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

	const first = await saved(viewer);
	for (const name of ['CFF Container', 'Swimlane List', 'Phase List', 'Swimlane', 'Separator'])
		expect(first.names).toContain(`${name} (hidden)`);
	expect(first.instances.get('50')).toBe('CFF Container (hidden)');
	expect(first.instances.get('51')).toBe('Swimlane List (hidden)');
	expect(first.instances.get('52')).toBe('Phase List (hidden)');
	for (const lane of ['10', '11', '12'])
		expect(first.instances.get(lane)).toBe('Swimlane (hidden)');
	// The page is related to the seven masters it uses (no separator yet), or Visio would not load
	// the instances.
	expect(first.rels.match(/relationships\/master"/g)!.length).toBe(7);

	// Insert Swimlane After and Separator through the ribbon: the new parts are instances too.
	await viewer
		.locator('svg.paper [data-shape-id="11"]')
		.click({ force: true, position: { x: 10, y: 10 } });
	await viewer.locator('office-ui-ribbon #tab-cross-functional').click();
	const panel = viewer.locator(TAB);
	await panel.getByRole('button', { name: 'Swimlane', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Insert Swimlane After"]`).click();
	await panel.getByRole('button', { name: 'Separator', exact: true }).click();
	await expect
		.poll(async () => {
			// Export is refused while an edit is still applying; poll again.
			const now = await saved(viewer).catch(() => undefined);
			if (!now) return undefined;
			const values = [...now.instances.values()];
			return [
				values.filter((name) => name === 'Swimlane (hidden)').length,
				values.filter((name) => name === 'Separator (hidden)').length,
			];
		})
		.toEqual([4, 1]);
	expect(errors).toEqual([]);
});
