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
	// Mostly attributes, as real shapes are: the page passes the node limit inside the heap budget.
	const shape = (id: number) =>
		`<Shape ID="${id}" Type="Shape">${'<C a="" b="" c="" d=""/>'.repeat(16)}</Shape>`;
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
		'Die Seite enthält 19000 Formen, mehr als dieser Viewer öffnen kann (er liest höchstens 1500000 XML-Knoten pro Teil).';
	await expect(viewer.element.load(await oversizedPage(19_000))).rejects.toThrow(refusal);
	expect(viewer.element.shadowRoot!.querySelector('[data-status]')!.textContent).toBe(refusal);
	viewer.destroy();
}, 60_000);

/** A page of `count` shapes of 18 ShapeSheet cells each, as Visio saves a Basic Shapes rectangle. */
async function cellHeavyPage(count: number): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await createVsdxFixture('Movable'));
	const cells = Array.from({ length: 18 }, (_, i) => `<Cell N="User.c${i}" V="${i}"/>`).join('');
	const shapes = Array.from(
		{ length: count },
		(_, i) => `<Shape ID="${i + 1}" Type="Shape">${cells}</Shape>`,
	).join('');
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace(/<Shapes>.*<\/Shapes>/s, `<Shapes>${shapes}</Shapes>`),
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'STORE' });
}

it('opens a drawing too large to edit read-only and says why in the viewer language', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const viewer = mountViewer(host, { document: demoDocument });
	viewer.element.setAttribute('locale', 'de');
	await viewer.element.load(await cellHeavyPage(4_000));
	expect(viewer.element.canEdit).toBe(false);
	expect(viewer.element.shadowRoot!.querySelector('[data-status]')!.textContent).toBe(
		'Diese Zeichnung ist zu groß, um sie hier zu bearbeiten: Ihre Seiten enthalten etwa 104000 ShapeSheet-Zellen, mehr als die 100000, denen eine Bearbeitung folgen kann. Sie wird schreibgeschützt geöffnet.',
	);
	// A drawing within the limit is editable again.
	await viewer.element.load(await cellHeavyPage(100));
	expect(viewer.element.canEdit).toBe(true);
	viewer.destroy();
}, 60_000);
