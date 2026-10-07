import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from 'ooxml-core/visio';
import { assertViewableDocument, copySnapshotScene } from 'ooxml-core/visio/ui';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { exportPageSvg } from './export-svg';
import { renderPage } from './render-svg';
import { createPrintSnapshot } from './print-snapshot';

it('shares normalized radial stops between live, export and immutable print snapshots', async () => {
	const document = await parseVsdx(await createVsdxFixture('Radial'));
	const paint = document.pages[0]!.shapes[0]!.style;
	paint.fillOpacity = 1;
	paint.fillGradient = {
		type: 'radial',
		center: [0, 1],
		radius: 1.4,
		stops: [
			{ offset: 0, color: '#ff0000', opacity: 0.8 },
			{ offset: 1, color: '#0000ff', opacity: 0.5 },
		],
	};
	const live = renderPage(document, document.pages[0]!);
	const gradient = live.svg.querySelector('radialGradient')!;
	expect(gradient.getAttribute('gradientUnits')).toBe('objectBoundingBox');
	expect(gradient.getAttribute('cx')).toBe('0');
	expect(gradient.getAttribute('cy')).toBe('1');
	expect(gradient.getAttribute('r')).toBe('1.4');
	expect(
		[...gradient.querySelectorAll('stop')].map((stop) => stop.getAttribute('stop-opacity')),
	).toEqual(['0.8', '0.5']);
	const exported = exportPageSvg(document).svg;
	const snapshot = createPrintSnapshot(document, { pageIndices: [0] });
	paint.fillGradient.radius = 0.73;
	for (const source of [exported, snapshot.pages[0]!.svg]) {
		expect(source).toContain('radialGradient');
		expect(source).toContain('r="1.4"');
		expect(source).toContain('stop-opacity="0.8"');
	}
	live.dispose();
});

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
	expect(live.svg.querySelector('path')!.getAttribute('shape-rendering')).toBe('crispEdges');
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
