// @vitest-environment jsdom
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { DocxEmbeddedFonts } from './embedded-fonts';
import { DocxWebfonts } from './docx-webfonts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const KEY = '{3A5C1F27-9B4E-4D80-A1C6-72E5D0B8F419}';

/** ISO/IEC 29500-2 14.2.1: the first 32 bytes XORed with the GUID's bytes, last byte first. */
function obfuscate(font: Uint8Array): Uint8Array {
	const hex = KEY.replace(/[{}-]/g, '');
	const key = Array.from({ length: 16 }, (_, i) =>
		parseInt(hex.slice(i * 2, i * 2 + 2), 16),
	).reverse();
	const out = font.slice();
	for (let i = 0; i < 32; i++) out[i] = out[i]! ^ key[i % 16]!;
	return out;
}

/** A stand-in for a TrueType file: the browser's `FontFace` is replaced, so only the bytes matter. */
const FONT = Uint8Array.from({ length: 64 }, (_, i) => (i < 4 ? [0, 1, 0, 0][i]! : i));

async function chivo() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Chivo"/></w:rPr><w:t>Hi</w:t></w:r></w:p></w:body></w:document>`,
	);
	zip.file(
		'word/fontTable.xml',
		`<w:fonts xmlns:w="${W}" xmlns:r="${R}"><w:font w:name="Chivo"><w:embedRegular r:id="rIdF1" w:fontKey="${KEY}"/><w:embedItalic r:id="rIdF2" w:fontKey="${KEY}"/></w:font></w:fonts>`,
	);
	zip.file(
		'word/_rels/fontTable.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdF1" Type="${R}/font" Target="fonts/font1.odttf"/><Relationship Id="rIdF2" Type="${R}/font" Target="fonts/font2.odttf"/></Relationships>`,
	);
	zip.file('word/fonts/font1.odttf', obfuscate(FONT));
	zip.file('word/fonts/font2.odttf', obfuscate(FONT));
	return (await loadDocx(await zip.generateAsync({ type: 'uint8array' }))).model;
}

class FakeFace {
	loaded = false;
	constructor(
		readonly family: string,
		readonly source: ArrayBuffer,
		readonly descriptors: { weight: string; style: string },
	) {}
	async load() {
		this.loaded = true;
		return this;
	}
}

function fakeSet() {
	const faces = new Set<FakeFace>();
	return {
		faces,
		add: (face: unknown) => faces.add(face as FakeFace),
		delete: (face: unknown) => faces.delete(face as FakeFace),
	};
}

describe('fonts embedded in a Word document', () => {
	afterEach(() => vi.unstubAllGlobals());

	it('registers the de-obfuscated faces for the open document and removes them with it', async () => {
		vi.stubGlobal('FontFace', FakeFace);
		const fonts = new DocxEmbeddedFonts();
		const set = fakeSet();
		const families = await fonts.load(document, await chivo(), set);
		expect(families).toEqual(['Chivo']);
		const faces = [...set.faces];
		expect(faces.map((face) => [face.family, face.descriptors])).toEqual([
			['Chivo', { weight: '400', style: 'normal' }],
			['Chivo', { weight: '400', style: 'italic' }],
		]);
		// The bytes are the font, not the obfuscated part.
		expect([...new Uint8Array(faces[0]!.source)]).toEqual([...FONT]);
		expect(faces.every((face) => face.loaded)).toBe(true);
		fonts.dispose();
		expect(set.faces.size).toBe(0);
		expect(fonts.families).toEqual([]);
	});

	it('does nothing for a document without embedded fonts, and without FontFace', async () => {
		const set = fakeSet();
		const fonts = new DocxEmbeddedFonts();
		const model = await chivo();
		expect(await fonts.load(document, model, set)).toEqual([]);
		expect(set.faces.size).toBe(0);
	});

	it('falls back quietly when the page has no font set', async () => {
		vi.stubGlobal('FontFace', FakeFace);
		const webfonts = new DocxWebfonts();
		await webfonts.prepare(document, await chivo());
		// Registered in the page's font set (jsdom has none): no throw, nothing requested.
		expect(webfonts.embedded.families).toEqual([]);
	});
});
