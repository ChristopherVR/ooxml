// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import JSZip from 'jszip';
import { demoDocument } from 'ooxml-core/visio/ui';
import { mountViewer } from './binding';
import { createVsdxFixture } from './__fixtures__/fixture.mjs';

afterEach(() => document.body.replaceChildren());

/** A page of `count` shapes past the XML node limit, stored uncompressed (no ratio refusal). */
async function oversizedPage(count: number): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Movable'));
	const shape = (id: number) => `<Shape ID="${id}" Type="Shape">${'<C/>'.repeat(80)}</Shape>`;
	const shapes = Array.from({ length: count }, (_, i) => shape(i + 1)).join('');
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shapes>.*<\/Shapes>/s, `<Shapes>${shapes}</Shapes>`),
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
}

it('refuses a page with too many shapes in the viewer language, with its shape count', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	viewer.element.setAttribute('locale', 'de');
	const refusal =
		'Die Seite enthält 16000 Formen, mehr als dieser Viewer öffnen kann (er liest höchstens 1250000 XML-Knoten pro Teil).';
	await expect(viewer.element.load(await oversizedPage(16_000))).rejects.toThrow(refusal);
	expect(viewer.element.shadowRoot!.querySelector('[data-status]')!.textContent).toBe(refusal);
	viewer.destroy();
}, 60_000);
