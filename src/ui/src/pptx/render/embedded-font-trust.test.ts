import type { PptxEmbeddedFont, PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { beforeEach, describe, expect, it } from 'vitest';

import { readFontFileStyle, selectTrustedEmbeddedFonts } from './embedded-font-trust';
import { buildEmbeddedFontStyles } from './embedded-fonts';
import { resetGoogleWebfontSessionCache, resolveGoogleWebfontHref } from './google-webfonts';
import { GlyphOutlineFontCache } from './text-warp-outline-font-cache';

/**
 * Bytes of a minimal sfnt carrying only the two tables the trust check reads:
 * `OS/2` (usWeightClass at +4, fsSelection at +62) and `head` (macStyle at +44).
 */
function sfnt(
	weight: number,
	options: { italic?: boolean; macItalicOnly?: boolean } = {},
): Uint8Array {
	const os2Length = 96;
	const headLength = 54;
	const os2Offset = 12 + 2 * 16;
	const headOffset = os2Offset + os2Length;
	const bytes = new Uint8Array(headOffset + headLength);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, 0x00010000);
	view.setUint16(4, 2);
	const record = (index: number, tag: string, offset: number, length: number) => {
		const at = 12 + index * 16;
		for (let i = 0; i < 4; i++) view.setUint8(at + i, tag.charCodeAt(i));
		view.setUint32(at + 8, offset);
		view.setUint32(at + 12, length);
	};
	record(0, 'OS/2', os2Offset, os2Length);
	record(1, 'head', headOffset, headLength);
	view.setUint16(os2Offset + 4, weight);
	const italicBit = options.italic && !options.macItalicOnly ? 0x1 : 0;
	view.setUint16(os2Offset + 62, italicBit || 0x40);
	view.setUint16(headOffset + 44, options.italic && options.macItalicOnly ? 0x2 : 0);
	return bytes;
}

function dataUrl(bytes: Uint8Array): string {
	let binary = '';
	for (const byte of bytes) binary += String.fromCharCode(byte);
	return `data:font/ttf;base64,${btoa(binary)}`;
}

function face(
	name: string,
	weight: number,
	slot: { bold?: boolean; italic?: boolean } = {},
	fileItalic = Boolean(slot.italic),
): PptxEmbeddedFont {
	return {
		name,
		dataUrl: dataUrl(sfnt(weight, { italic: fileItalic })),
		format: 'truetype',
		...slot,
	};
}

/** A family's four slots holding the given file weights. */
function family(name: string, regular: number, bold = 700, italic = 400, boldItalic = 700) {
	return [
		face(name, regular),
		face(name, bold, { bold: true }),
		face(name, italic, { italic: true }),
		face(name, boldItalic, { bold: true, italic: true }),
	];
}

function deckUsing(...families: string[]): PptxSlide[] {
	const elements = families.map(
		(fontFamily) =>
			({
				type: 'text',
				textSegments: [{ text: 'x', style: { fontFamily } }],
			}) as unknown as PptxElement,
	);
	return [{ elements } as unknown as PptxSlide];
}

beforeEach(() => {
	resetGoogleWebfontSessionCache();
});

describe('readFontFileStyle', () => {
	it('reads the OS/2 weight and the italic flags', () => {
		expect(readFontFileStyle(sfnt(300))).toEqual({ weight: 300, italic: false });
		expect(readFontFileStyle(sfnt(400, { italic: true }))).toEqual({ weight: 400, italic: true });
		expect(readFontFileStyle(sfnt(400, { italic: true, macItalicOnly: true }))).toEqual({
			weight: 400,
			italic: true,
		});
	});

	it('scales the old 1-9 weight classes to CSS weights', () => {
		expect(readFontFileStyle(sfnt(7))?.weight).toBe(700);
	});

	it('leaves files it cannot read unjudged', () => {
		expect(
			readFontFileStyle(new Uint8Array([0x77, 0x4f, 0x46, 0x32, 0, 0, 0, 0, 0, 0, 0, 0])),
		).toBeUndefined();
		expect(readFontFileStyle(sfnt(0))).toBeUndefined();
		expect(readFontFileStyle(new Uint8Array(4))).toBeUndefined();
	});
});

describe('selectTrustedEmbeddedFonts', () => {
	it('keeps a Google Fonts family whose files match their slots', () => {
		const fonts = family('DM Sans', 400);
		expect(selectTrustedEmbeddedFonts(fonts)).toEqual(fonts);
	});

	it('rejects the whole typeface when one slot holds another weight', () => {
		// The regular slot is a SemiBold file: every regular run would paint SemiBold.
		expect(selectTrustedEmbeddedFonts(family('DM Sans', 600))).toEqual([]);
	});

	it('judges weight-named typefaces against the weight their name implies', () => {
		const light = family('DM Sans Light', 300, 700, 300, 700);
		expect(selectTrustedEmbeddedFonts(light)).toEqual(light);
		// The same files a mislabelled deck reuses for every weight.
		expect(selectTrustedEmbeddedFonts(family('DM Sans Light', 600))).toEqual([]);
		expect(selectTrustedEmbeddedFonts(family('DM Sans Medium', 600, 700, 400))).toEqual([]);
		// Bold in a Black family is still Black (900), not the family's Bold file.
		expect(selectTrustedEmbeddedFonts([face('DM Sans Black', 700, { bold: true })])).toEqual([]);
	});

	it('rejects a slot whose file has the other style', () => {
		expect(selectTrustedEmbeddedFonts([face('DM Sans', 400, { italic: true }, false)])).toEqual([]);
	});

	it('tolerates foundry rounding below a full weight step', () => {
		const book = [face('Roboto', 350), face('Roboto', 750, { bold: true })];
		expect(selectTrustedEmbeddedFonts(book)).toEqual(book);
	});

	it('keeps brand fonts Google does not serve, whatever their files are', () => {
		const brand = family('Contoso Sans Light', 600);
		expect(selectTrustedEmbeddedFonts(brand)).toEqual(brand);
	});

	it('keeps a catalogue family named with a weight word at its own regular weight', () => {
		// "Archivo Black" is its own single-weight family on Google Fonts.
		const fonts = [face('Archivo Black', 400)];
		expect(selectTrustedEmbeddedFonts(fonts)).toEqual(fonts);
	});

	it('keeps entries whose bytes cannot be read', () => {
		const unreadable = [
			{ name: 'DM Sans', dataUrl: 'data:font/ttf;base64,AAEAAA==', format: 'truetype' },
		];
		expect(selectTrustedEmbeddedFonts(unreadable)).toEqual(unreadable);
	});

	it('judges only the rejected typeface, not its neighbours', () => {
		const good = family('Roboto Mono', 400);
		const trusted = selectTrustedEmbeddedFonts([...family('DM Sans', 600), ...good]);
		expect(trusted).toEqual(good);
	});

	it('returns the same answer for the same list without re-reading it', () => {
		const fonts = family('DM Sans', 600);
		expect(selectTrustedEmbeddedFonts(fonts)).toBe(selectTrustedEmbeddedFonts(fonts));
		expect(selectTrustedEmbeddedFonts(undefined)).toEqual([]);
	});
});

describe('embedded fonts that do not match their typeface', () => {
	it('are not registered as @font-face, but are still listed as embedded', () => {
		const result = buildEmbeddedFontStyles(
			[...family('DM Sans Light', 600), ...family('Roboto Mono', 400)],
			() => null,
		);
		expect(result.fontFaceCss).not.toContain('"DM Sans Light"');
		expect(result.fontFaceCss.match(/@font-face/g)).toHaveLength(4);
		expect(result.fontFamilies.some((list) => list.startsWith('"DM Sans Light"'))).toBe(true);
	});

	it('are requested from Google Fonts instead, at the weight the name implies', async () => {
		const href = await resolveGoogleWebfontHref(
			deckUsing('DM Sans Light', 'Roboto Mono'),
			[...family('DM Sans Light', 600), ...family('Roboto Mono', 400)],
			() => false,
		);
		expect(href).toContain(
			`family=${encodeURIComponent('DM Sans:ital,wght@0,300;0,400;0,700;1,300;1,400;1,700')}`,
		);
		// A matching embedded family is still served from the deck.
		expect(href).not.toContain('Roboto%20Mono');
	});

	it('are not parsed for glyph outlines', () => {
		const cache = new GlyphOutlineFontCache();
		cache.registerEmbeddedFonts(family('DM Sans', 600));
		expect(cache.hasAttempted('DM Sans', false, false)).toBe(false);
	});
});

describe('embedded weight-named faces that match', () => {
	it('register at the weight their text renders at', () => {
		const result = buildEmbeddedFontStyles(family('DM Sans Light', 300, 700, 300, 700), () => null);
		const weights = [...result.fontFaceCss.matchAll(/font-weight: (\d+)/g)].map((m) => m[1]);
		expect(weights).toEqual(['300', '700', '300', '700']);
	});
});
