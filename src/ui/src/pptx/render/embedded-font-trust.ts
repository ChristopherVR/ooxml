/**
 * embedded-font-trust.ts: decide whether a deck's embedded font files really
 * are the faces they are registered as.
 *
 * An embedded font is registered under the deck's typeface name, with its
 * weight and style taken from the slot it sits in (regular / bold / italic /
 * boldItalic). Nothing in the package guarantees the file matches the slot:
 * authoring tools have been seen to write one family's files into every
 * weight-named typeface ("DM Sans Light", "DM Sans Medium", ... all pointing at
 * the same SemiBold file). Registered blindly, every Light run renders
 * SemiBold.
 *
 * The check compares each file's own OS/2 `usWeightClass` and italic flags
 * with what its slot and typeface name imply (see `resolveFontWeight`). It
 * only judges typefaces that Google Fonts serves, the one case where a
 * correct replacement is available: a rejected typeface is requested from the
 * catalogue instead. A brand font with no other source keeps its embedded
 * files whatever they contain, since a wrong weight of the right typeface
 * still beats a generic fallback.
 *
 * A typeface is accepted or rejected as a whole. Keeping some of its faces
 * would be worse than keeping none: CSS matches weights within the first
 * family that has any face, so a lone bold face would paint the regular text
 * bold.
 *
 * @module render/embedded-font-trust
 */
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';

import { resolveEmbeddedFontClearBytes } from './embedded-font-bytes';
import { resolveFontItalic, resolveFontWeight, resolveGoogleFontsFace } from './font-weight-names';

/** Weight and style a font file declares about itself. */
export interface FontFileStyle {
	/** OS/2 `usWeightClass`, normalised to the CSS 1-1000 scale. */
	weight: number;
	/** OS/2 `fsSelection` ITALIC/OBLIQUE or `head.macStyle` italic. */
	italic: boolean;
}

/**
 * A difference of a full weight step (100) means a different face: SemiBold
 * (600) is not Medium (500). Smaller differences are foundry rounding
 * ("Book" at 350 or 380 for a regular slot).
 */
const WEIGHT_TOLERANCE = 100;

/** sfnt versions this reader understands: TrueType, OpenType CFF, Apple 'true'. */
const SFNT_VERSIONS = new Set([0x00010000, 0x4f54544f, 0x74727565]);

/**
 * Read the weight and italic flags from an uncompressed sfnt (TTF/OTF) font,
 * or `undefined` when the bytes are not one or lack an OS/2 table. WOFF/WOFF2
 * files compress their tables and are left unjudged.
 */
export function readFontFileStyle(bytes: Uint8Array): FontFileStyle | undefined {
	if (bytes.length < 12) {
		return undefined;
	}
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (!SFNT_VERSIONS.has(view.getUint32(0))) {
		return undefined;
	}
	const numTables = view.getUint16(4);
	let os2: number | undefined;
	let head: number | undefined;
	for (let i = 0; i < numTables; i++) {
		const record = 12 + i * 16;
		if (record + 16 > bytes.length) {
			return undefined;
		}
		const tag = view.getUint32(record);
		const offset = view.getUint32(record + 8);
		if (tag === 0x4f532f32) {
			os2 = offset; // 'OS/2'
		} else if (tag === 0x68656164) {
			head = offset; // 'head'
		}
	}
	if (os2 === undefined || os2 + 64 > bytes.length) {
		return undefined;
	}
	const weightClass = view.getUint16(os2 + 4);
	if (weightClass === 0 || weightClass > 1000) {
		return undefined;
	}
	const fsSelection = view.getUint16(os2 + 62);
	const macItalic =
		head !== undefined && head + 46 <= bytes.length && (view.getUint16(head + 44) & 0x2) !== 0;
	return {
		// Some old fonts use the 1-9 scale of the original OS/2 spec.
		weight: weightClass < 10 ? weightClass * 100 : weightClass,
		italic: (fsSelection & 0x201) !== 0 || macItalic,
	};
}

/** Per-list memo: bindings call this from several effects for the same deck. */
const trustedByList = new WeakMap<readonly PptxEmbeddedFont[], PptxEmbeddedFont[]>();

/** Per-list, per-typeface memo of {@link isEmbeddedTypefaceTrusted}. */
const typefaceTrustByList = new WeakMap<readonly PptxEmbeddedFont[], Map<string, boolean>>();

function trimmedName(font: PptxEmbeddedFont): string {
	return typeof font.name === 'string' ? font.name.trim() : '';
}

/**
 * Whether the embedded files of `typeface` in `fonts` may be registered (see
 * the module comment). Reads only that typeface's files, so a caller that
 * needs one family does not decode the others; the answer is the same as
 * membership in {@link selectTrustedEmbeddedFonts}.
 */
export function isEmbeddedTypefaceTrusted(
	fonts: readonly PptxEmbeddedFont[],
	typeface: string,
): boolean {
	const name = typeface.trim();
	let memo = typefaceTrustByList.get(fonts);
	const known = memo?.get(name);
	if (known !== undefined) {
		return known;
	}
	let trusted = true;
	if (name && resolveGoogleFontsFace(name)) {
		for (const font of fonts) {
			if (trimmedName(font) !== name) {
				continue;
			}
			const bytes = resolveEmbeddedFontClearBytes(font);
			const actual = bytes ? readFontFileStyle(bytes) : undefined;
			if (!actual) {
				continue;
			}
			// A style variant named as its own family ("Lato Light Italic") holds an
			// italic file in its regular slot.
			const expectedWeight = resolveFontWeight(name, Boolean(font.bold));
			const expectedItalic = resolveFontItalic(name, font.italic);
			if (
				Math.abs(actual.weight - expectedWeight) >= WEIGHT_TOLERANCE ||
				actual.italic !== expectedItalic
			) {
				trusted = false;
				break;
			}
		}
	}
	if (!memo) {
		memo = new Map();
		typefaceTrustByList.set(fonts, memo);
	}
	memo.set(name, trusted);
	return trusted;
}

/**
 * The embedded fonts that may be registered as `@font-face` rules, and that
 * count as covering their family when deciding what to request from Google
 * Fonts. See the module comment for the rule.
 */
export function selectTrustedEmbeddedFonts(
	fonts: readonly PptxEmbeddedFont[] | null | undefined,
): PptxEmbeddedFont[] {
	if (!fonts || fonts.length === 0) {
		return [];
	}
	const memo = trustedByList.get(fonts);
	if (memo) {
		return memo;
	}
	const trusted = fonts.filter((font) => isEmbeddedTypefaceTrusted(fonts, trimmedName(font)));
	trustedByList.set(fonts, trusted);
	return trusted;
}
