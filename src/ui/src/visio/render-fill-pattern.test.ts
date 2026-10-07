import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import { assertViewableDocument, copySnapshotScene } from 'ooxml-core/visio/ui';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { exportPageSvg } from './export-svg';
import { renderPage } from './render-svg';

it('shares bounded pattern resources between live rendering, portable exports and snapshots', async () => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Pattern'));
	const page = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		page.replace(
			'<Text>',
			'<Cell N="FillPattern" V="2"/><Cell N="FillForegnd" V="#ff0000"/><Cell N="FillBkgnd" V="#0000ff"/><Text>',
		),
	);
	const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
	const tile = document.pages[0]!.shapes[0]!.style.fillPattern!;
	expect(tile).toBeDefined();
	const live = renderPage(document, document.pages[0]!);
	expect(live.svg.querySelector('pattern image')!.getAttribute('href')).toMatch(/^blob:/);
	const output = exportPageSvg(document);
	expect(output.svg).toContain('data:image/png;base64,');
	expect(output.svg).toContain('patternUnits="userSpaceOnUse"');
	expect(output.svg).not.toContain('blob:');
	const snapshot = copySnapshotScene(document);
	const copied = snapshot.pages[0]!.shapes[0]!.style.fillPattern!;
	expect(copied.bytes).toEqual(tile.bytes);
	expect(copied.bytes).not.toBe(tile.bytes);
	tile.width = 0;
	expect(() => assertViewableDocument(document)).toThrow('fill pattern width');
	tile.width = 1 / 12;
	tile.bytes[0] = 0;
	expect(() => assertViewableDocument(document)).toThrow();
	assertViewableDocument(snapshot);
	live.dispose();
});
