import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

const NS = 'http://schemas.microsoft.com/office/visio/2012/main';
const REL = 'http://schemas.microsoft.com/visio/2010/relationships/';
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const cell = (name: string, value: number | string, formula = '') =>
	`<Cell N="${name}" V="${value}"${formula ? ` F="${formula}"` : ''}/>`;
const box = `${cell('PinX', 0.5)}${cell('PinY', 0.5)}${cell('Width', 1)}${cell('Height', 0.75)}${cell('LocPinX', 0.5, 'Width*0.5')}${cell('LocPinY', 0.375, 'Height*0.5')}`;
const outline = `<Section N="Geometry" IX="0"><Row T="RelMoveTo" IX="1">${cell('X', 0)}${cell('Y', 0)}</Row><Row T="RelLineTo" IX="2">${cell('X', 1)}${cell('Y', 0)}</Row><Row T="RelLineTo" IX="3">${cell('X', 1)}${cell('Y', 1)}</Row><Row T="RelLineTo" IX="4">${cell('X', 0)}${cell('Y', 1)}</Row><Row T="RelLineTo" IX="5">${cell('X', 0)}${cell('Y', 0)}</Row></Section>`;
const points = `<Section N="Connection"><Row T="Connection" IX="0">${cell('X', 0.5, 'Width*0.5')}${cell('Y', 0.75, 'Height*1')}</Row></Section>`;
const data = `<Section N="Property"><Row N="Pressure">${cell('Label', 'Pressure')}${cell('Value', 'bar')}${cell('Type', 0)}</Row></Section>`;

/** A .vssx: a package with the stencil main part and one master with a connection point and data. */
async function stencil(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await createVsdxFixture('unused'));
	const types = await zip.file('[Content_Types].xml')!.async('string');
	zip.file('[Content_Types].xml', types.replace('drawing.main', 'stencil.main'));
	zip.file(
		'visio/_rels/document.xml.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}pages" Target="pages/pages.xml"/><Relationship Id="rId2" Type="${REL}masters" Target="masters/masters.xml"/></Relationships>`,
	);
	zip.file(
		'visio/masters/masters.xml',
		`<Masters xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><Master ID="2" NameU="Valve" Name="Valve" UniqueID="{AAAAAAAA-0000-0000-0000-000000000009}"><PageSheet>${cell('PageWidth', 1)}${cell('PageHeight', 0.75)}</PageSheet><Rel r:id="rId1"/></Master></Masters>`,
	);
	zip.file(
		'visio/masters/_rels/masters.xml.rels',
		`<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}master" Target="master1.xml"/></Relationships>`,
	);
	zip.file(
		'visio/masters/master1.xml',
		`<MasterContents xmlns="${NS}"><Shapes><Shape ID="5" Type="Shape">${box}${points}${data}${outline}</Shape></Shapes></MasterContents>`,
	);
	return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

test('opens a stencil file, drops its master and reopens the saved drawing', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1440, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'blank.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await createVsdxFixture('Existing shape'),
	});
	await expect(page.locator('#file-name')).toHaveText('blank.vsdx');
	await viewer.locator('input[data-open-stencil]').setInputFiles({
		name: 'Valves.vssx',
		mimeType: 'application/vnd.ms-visio.stencil',
		buffer: await stencil(),
	});
	const section = viewer.locator('[data-stencil="file-1"]');
	await expect(section.locator('.stencil-title')).toHaveText('Valves');
	const valve = section.locator('[data-master="file:1:2"]');
	await expect(valve).toHaveText('Valve');
	await expect(valve.locator('svg.master-preview path').first()).toBeAttached();
	const shapes = viewer.locator('svg.paper > g > [data-shape-id]');
	const before = await shapes.count();
	await valve.dragTo(viewer.locator('svg.paper'), { targetPosition: { x: 200, y: 200 } });
	await expect(viewer.locator('[data-status]')).toHaveText(
		/Valve .+ added from the opened stencil/,
	);
	await expect(shapes).toHaveCount(before + 1);
	// My Shapes keeps a collected master for the session.
	await valve.click({ button: 'right' });
	await viewer.locator('[data-shapes-menu="master"] [command="my-shapes"]').click();
	await expect(viewer.locator('[data-stencil="my-shapes"] [data-master]')).toHaveCount(1);
	// The saved drawing carries the master, and reopening it resolves the instance.
	const saved = await viewer.evaluate((node) => [
		...(node as VisioViewerElement).controller.exportVsdx().bytes,
	]);
	await page.locator('#file').setInputFiles({
		name: 'with-valve.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: Buffer.from(saved),
	});
	await expect(page.locator('#file-name')).toHaveText('with-valve.vsdx');
	const reopened = await viewer.evaluate((node) => {
		const model = (node as VisioViewerElement).controller.state.document!;
		const shape = model.pages[0]!.shapes.at(-1)!;
		return {
			master: model.masters!.find((master) => master.id === shape.masterId)?.name,
			points: shape.connectionPoints?.length,
			data: shape.shapeData?.map((item) => item.label),
		};
	});
	expect(reopened).toEqual({ master: 'Valve', points: 1, data: ['Pressure'] });
	await expect(viewer.locator('[data-stencil="document"] [data-master^="document:"]')).toHaveText(
		'Valve',
	);
	expect(errors).toEqual([]);
});
