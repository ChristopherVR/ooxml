import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx } from 'ooxml-core/visio';
import { assertViewableDocument, copySnapshotScene } from 'ooxml-core/visio/ui';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';
import { renderPage } from './render-svg';
import reference from '../../../core/visio/__fixtures__/layer-colors-native.json';

it('paints colored-layer geometry and text, and retains alpha in immutable snapshots', async () => {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Layer color'));
	const pages = await zip.file('visio/pages/pages.xml')!.async('string');
	zip.file(
		'visio/pages/pages.xml',
		pages.replace(
			'</PageSheet>',
			'<Section N="Layer"><Row IX="0"><Cell N="Color" V="#ff0000"/><Cell N="ColorTrans" V="0.4"/></Row></Section></PageSheet>',
		),
	);
	const contents = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		contents.replace(
			'<Text>',
			'<Cell N="LayerMember" V="0"/><Section N="Character"><Row IX="0"><Cell N="Color" V="#00ffff"/><Cell N="ColorTrans" V="0.5"/></Row></Section><Text>',
		),
	);
	const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
	const result = renderPage(document, document.pages[0]!);
	const path = result.svg.querySelector('[data-shape-id="1"] path')!;
	expect(path.getAttribute('fill')).toBe('#ffffff');
	expect(path.getAttribute('stroke')).toBe('#ff0000');
	expect(path.getAttribute('fill-opacity')).toBe('0.6');
	expect(path.getAttribute('stroke-opacity')).toBe('0.6');
	const run = result.svg.querySelector('text tspan tspan')!;
	expect(run.getAttribute('fill')).toBe('#ff0000');
	expect(run.getAttribute('fill-opacity')).toBe('0.6');
	const snapshot = copySnapshotScene(document);
	expect(snapshot.pages[0]!.shapes[0]!.text.opacity).toBe(0.6);
	expect(snapshot.pages[0]!.shapes[0]!.text.runs[0]!.opacity).toBe(0.6);
	document.pages[0]!.shapes[0]!.text.runs[0]!.opacity = 0;
	expect(snapshot.pages[0]!.shapes[0]!.text.runs[0]!.opacity).toBe(0.6);
	document.pages[0]!.shapes[0]!.text.opacity = NaN;
	expect(() => assertViewableDocument(document)).toThrow('text opacity');
	document.pages[0]!.shapes[0]!.text.opacity = 0.6;
	document.pages[0]!.shapes[0]!.text.runs[0]!.opacity = 2;
	expect(() => assertViewableDocument(document)).toThrow('text run opacity');
	result.dispose();
});

const nativeDirectory = process.env.VISIO_NATIVE_LAYER_COLORS_DIR;
it.skipIf(!nativeDirectory)(
	'paints every native page with the captured geometry and text alpha',
	async () => {
		const document = await parseVsdx(
			new Uint8Array(await readFile(join(nativeDirectory!, 'layer-colors.vsdx'))),
		);
		for (let index = 0; index < document.pages.length; index++) {
			const result = renderPage(document, document.pages[index]!);
			const actual = result.svg.querySelector('[data-shape-id="1"] path')!;
			const native = reference.cases[index]!.paint;
			expect(actual.getAttribute('stroke')).toBe(native.stroke);
			expect(
				Math.abs(Number(actual.getAttribute('stroke-opacity')) - native.strokeOpacity),
			).toBeLessThan(0.0050001);
			const run = result.svg.querySelector('text tspan tspan')!;
			expect(run.getAttribute('fill')).toBe(native.text);
			expect(Number(run.getAttribute('fill-opacity'))).toBeCloseTo(native.textOpacity, 6);
			result.dispose();
		}
	},
);
