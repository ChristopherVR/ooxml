/**
 * `saveDocx` with an injected `now` and a seeded id source gives the same bytes whenever it runs:
 * ZIP entries are dated 1980-01-01 (as Word writes) and the ids Word draws at random (here the
 * bibliography data store item) come from the source. A new document names no face in its
 * styles; they refer to the theme fonts, which its theme part names.
 */
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_THEME_FONTS } from '../drawingml/office-theme-fonts';
import { seededIdSource } from '../opc/reproducible';
import { DEFAULT_STYLES_XML } from './default-styles';
import { createDocument, type DocumentModel } from './model';
import { saveDocx } from './save';

const NOW = new Date('2026-01-15T09:00:00Z');
const SOURCES =
	'<b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography" SelectedStyle="" StyleName=""/>';

function model(): DocumentModel {
	const doc = createDocument();
	doc.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Hello' }] }];
	doc.packageInformation = { bibliography: { xml: SOURCES } };
	return doc;
}

const save = () => saveDocx(model(), undefined, { now: NOW, ids: seededIdSource(7) });

afterEach(() => {
	vi.useRealTimers();
});

describe('reproducible Word saves', () => {
	it('gives identical bytes for the same now and seed at different times', async () => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-05-05T10:00:00Z'));
		const first = await save();
		vi.setSystemTime(new Date('2026-06-07T11:00:03Z'));
		const second = await save();
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
		const view = new DataView(first.buffer, first.byteOffset);
		expect(view.getUint16(12, true)).toBe(0x21);
	});

	it('draws the data store item id from the source, at random by default', async () => {
		const itemId = async (bytes: Uint8Array) => {
			const zip = await JSZip.loadAsync(bytes);
			const name = Object.keys(zip.files).find((part) =>
				/^customXml\/itemProps\d+\.xml$/u.test(part),
			);
			const xml = await zip.file(name!)!.async('string');
			return /ds:itemID="([^"]+)"/u.exec(xml)?.[1];
		};
		expect(await itemId(await save())).toBe(await itemId(await save()));
		const random = await itemId(await saveDocx(model()));
		expect(random).toMatch(/^\{[0-9A-F-]{36}\}$/u);
		expect(random).not.toBe(await itemId(await saveDocx(model())));
	});

	it('refers new styles to the theme fonts and saves the default theme', async () => {
		// Only w:rFonts names fonts; w:lang also has a w:eastAsia attribute (a language).
		expect(DEFAULT_STYLES_XML).not.toMatch(/<w:rFonts[^>]*\sw:(?:ascii|hAnsi|eastAsia|cs)="/u);
		expect(DEFAULT_STYLES_XML).toContain('w:asciiTheme="minorHAnsi"');
		expect(DEFAULT_STYLES_XML).toContain('w:asciiTheme="majorHAnsi"');
		const zip = await JSZip.loadAsync(await save());
		const theme = await zip.file('word/theme/theme1.xml')?.async('string');
		expect(theme).toContain(`<a:latin typeface="${OFFICE_THEME_FONTS.major}"`);
		expect(theme).toContain(`<a:latin typeface="${OFFICE_THEME_FONTS.minor}"`);
	});
});
