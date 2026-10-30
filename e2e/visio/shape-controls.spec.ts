import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { openDemo } from './demo-page';

/**
 * A synthetic shape (no Microsoft content) with one horizontal control handle that slants its top
 * edge, a hidden second geometry section a SETF action toggles, and an action Visio runs outside
 * the ShapeSheet.
 */
async function controlsFixture(text = 'Slanted', tabs = ''): Promise<Buffer> {
	const zip = new JSZip();
	const rels = 'http://schemas.openxmlformats.org/package/2006/relationships';
	const visio = 'http://schemas.microsoft.com/visio/2010/relationships/';
	const ns = 'http://schemas.microsoft.com/office/visio/2012/main';
	const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/visio/document.xml" ContentType="application/vnd.ms-visio.drawing.main+xml"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}document" Target="visio/document.xml"/></Relationships>`,
	);
	zip.file('visio/document.xml', `<VisioDocument xmlns="${ns}"/>`);
	zip.file(
		'visio/_rels/document.xml.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}pages" Target="pages/pages.xml"/></Relationships>`,
	);
	zip.file(
		'visio/pages/pages.xml',
		`<Pages xmlns="${ns}" xmlns:r="${relNs}"><Page ID="1" Name="Controls"><PageSheet><Cell N="PageWidth" V="8.5"/><Cell N="PageHeight" V="11"/></PageSheet><Rel r:id="rId1"/></Page></Pages>`,
	);
	zip.file(
		'visio/pages/_rels/pages.xml.rels',
		`<Relationships xmlns="${rels}"><Relationship Id="rId1" Type="${visio}page" Target="page1.xml"/></Relationships>`,
	);
	const cell = (name: string, value: string | number, formula = '') =>
		`<Cell N="${name}" V="${value}"${formula ? ` F="${formula}"` : ''}/>`;
	const row = (index: number, type: string, cells: string) =>
		`<Row T="${type}" IX="${index}">${cells}</Row>`;
	const shape =
		`<Shape ID="1" NameU="Slanted" Type="Shape">` +
		cell('Width', 3) +
		cell('Height', 1.5) +
		cell('PinX', 4) +
		cell('PinY', 7) +
		cell('LocPinX', 1.5, 'Width*0.5') +
		cell('LocPinY', 0.75, 'Height*0.5') +
		`<Section N="Control"><Row N="Row_1">${cell('X', 0.6, 'BOUND(Width*0.2,0,FALSE,0,Width)')}${cell('Y', 1.5, 'Height*1')}${cell('XCon', 0)}${cell('YCon', 1)}${cell('CanGlue', 0)}${cell('Prompt', 'Slant')}</Row></Section>` +
		`<Section N="Actions"><Row N="Row_1">${cell('Menu', 'Show Bo&amp;x', 'IF(Geometry2.NoShow=1,&quot;Show Bo&amp;x&quot;,&quot;Hide Bo&amp;x&quot;)')}${cell('Action', 0, 'SETF(GetRef(Geometry2.NoShow),NOT(Geometry2.NoShow))')}${cell('Disabled', 0)}</Row><Row N="Row_2">${cell('Menu', 'Open Help')}${cell('Action', 0, 'OPENFILE(&quot;help.chm&quot;)')}${cell('Disabled', 0)}</Row></Section>` +
		`<Section N="Geometry" IX="0">${row(1, 'MoveTo', cell('X', 0, 'Width*0') + cell('Y', 0, 'Height*0'))}${row(2, 'LineTo', cell('X', 3, 'Width*1') + cell('Y', 0, 'Height*0'))}${row(3, 'LineTo', cell('X', 3, 'Width*1') + cell('Y', 1.5, 'Height*1'))}${row(4, 'LineTo', cell('X', 0.6, 'Controls.Row_1') + cell('Y', 1.5, 'Height*1'))}${row(5, 'LineTo', cell('X', 0, 'Width*0') + cell('Y', 0, 'Height*0'))}</Section>` +
		`<Section N="Geometry" IX="1">${cell('NoFill', 1)}${cell('NoShow', 1)}${row(1, 'MoveTo', cell('X', 0.3) + cell('Y', 0.3))}${row(2, 'LineTo', cell('X', 1.2) + cell('Y', 0.3))}</Section>` +
		tabs +
		`<Text><tp IX="0"/>${text}</Text></Shape>`;
	zip.file(
		'visio/pages/page1.xml',
		`<PageContents xmlns="${ns}"><Shapes>${shape}</Shapes></PageContents>`,
	);
	return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

test('a control handle is dragged and a shape action runs from the context menu', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'controls.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await controlsFixture(),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Slanted');
	const status = viewer.locator('[data-status]');
	const shape = viewer.locator('svg.paper [data-shape-id="1"]');
	const geometry = () => shape.locator('[data-geometry]').first().getAttribute('d');

	await shape.click();
	const handle = viewer.locator('[data-control-handle="Row_1"]');
	await expect(handle).toBeVisible();
	const before = await geometry();
	const box = (await handle.boundingBox())!;
	await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
	await page.mouse.down();
	await page.mouse.move(box.x + box.width / 2 + 60, box.y + box.height / 2, { steps: 6 });
	await page.mouse.up();
	// The top edge's corner moved right: the geometry that reads the control was recomputed.
	await expect.poll(geometry).not.toBe(before);

	await shape.click({ button: 'right' });
	const menu = viewer.locator('[data-context-menu="shape"]');
	const hide = menu.locator('office-ui-menu-item[data-shape-action="Row_1"]');
	await expect(hide).toHaveAttribute('label', 'Show Box');
	const help = menu.locator('office-ui-menu-item[data-shape-action="Row_2"]');
	await expect(help).toHaveAttribute('disabled', '');
	await expect(help).toHaveAttribute('title', /OPENFILE/);
	const parts = await shape.locator('[data-geometry]').count();
	await hide.getByRole('menuitem').click();
	await expect(status).toHaveText('Show Box done.');
	await expect.poll(() => shape.locator('[data-geometry]').count()).toBeGreaterThan(parts);

	await shape.click({ button: 'right' });
	await expect(menu.locator('office-ui-menu-item[data-shape-action="Row_1"]')).toHaveAttribute(
		'label',
		'Hide Box',
	);
	expect(errors).toEqual([]);
});

test('tab stops of the Tabs section place tabbed text', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'tabs.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await controlsFixture(
			'Item\tQty\tPrice',
			'<Section N="Paragraph"><Row IX="0"><Cell N="HorzAlign" V="0"/></Row></Section><Section N="Tabs"><Row IX="0"><Cell N="Position1" V="1"/><Cell N="Alignment1" V="0"/><Cell N="Position2" V="2.5"/><Cell N="Alignment2" V="2"/></Row></Section>',
		),
	});
	await expect(viewer.locator('svg.paper')).toContainText('Item');
	const rows = viewer.locator('svg.paper [data-shape-id="1"] text > tspan[x]');
	await expect(rows).toHaveCount(3);
	const xs = await rows.evaluateAll((nodes) => nodes.map((node) => Number(node.getAttribute('x'))));
	// The text block starts at its left margin; Qty starts at the 1 inch stop and Price ends at 2.5 inches.
	expect(xs[1]! - xs[0]!).toBeCloseTo(1 - 0, 1);
	expect(xs[2]!).toBeGreaterThan(xs[1]! + 0.3);
	expect(xs[2]!).toBeLessThan(2.5 + xs[0]!);
});
