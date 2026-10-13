import { test, expect } from '@playwright/test';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { createVsdxFixture } from './fixture.mjs';
import { openDemo } from './demo-page';

test.use({ timezoneId: 'UTC' });
test('ISO date Field choice inserts a creation date and survives history and export', async ({
	page,
}) => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Created: '));
	const relationships = await zip.file('_rels/.rels')!.async('string');
	const types = await zip.file('[Content_Types].xml')!.async('string');
	zip.file(
		'[Content_Types].xml',
		types.replace(
			'</Types>',
			'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
		),
	);
	zip.file(
		'_rels/.rels',
		relationships.replace(
			'</Relationships>',
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
		),
	);
	zip.file(
		'docProps/core.xml',
		'<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/"><dcterms:created>2026-10-13T04:15:16Z</dcterms:created></cp:coreProperties>',
	);
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Date.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	await expect(page.locator('#file-name')).toHaveText('Date.vsdx');
	const paper = viewer.locator('svg.paper');
	await paper.locator('[data-shape-id="1"]').first().click();
	await viewer.getByRole('tab', { name: 'Insert', exact: true }).click();
	await viewer.locator('[command="field"] button').click();
	const field = viewer.locator('.field-dialog');
	await field.locator('[name="field-category"]').selectOption('date');
	await field.locator('[name="field-field"]').selectOption('DOCCREATION()');
	await field.locator('[name="field-format"]').selectOption({ label: 'ISO date (2026-10-09)' });
	await field.locator('[label="OK"] button').click();
	await expect(paper).toContainText('Created: 2026-10-13');
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
	await expect(paper).toContainText('Created:');
	await expect(paper).not.toContainText('2026-10-13');
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.redo());
	await expect(paper).toContainText('Created: 2026-10-13');
	const bytes = await viewer.evaluate((node) =>
		Array.from((node as VisioViewerElement).exportVsdx().bytes),
	);
	const text = (await parseVsdx(new Uint8Array(bytes))).pages[0]!.shapes[0]!.text;
	expect(text.plainText).toBe('Created: 2026-10-13');
	expect(text.fields).toMatchObject([{ formula: 'DOCCREATION()', cached: '2026-10-13' }]);
	await page.locator('#file').setInputFiles({
		name: 'Saved date.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: Buffer.from(bytes),
	});
	await expect(page.locator('#file-name')).toHaveText('Saved date.vsdx');
	await expect(paper).toContainText('Created: 2026-10-13');
});
