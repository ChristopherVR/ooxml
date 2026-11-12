import { GOOGLE_FONTS_FAMILIES, GOOGLE_FONTS_VARIANT_MASKS } from './google-fonts-catalogue';
import { findGoogleFontsFamily } from './google-fonts-lookup';

/** A supported standard-weight face in the bundled Google Fonts catalogue. */
export interface GoogleFontsVariant {
	readonly weight: number;
	readonly italic: boolean;
}

const EMPTY_VARIANTS: readonly GoogleFontsVariant[] = Object.freeze([]);
const variantsByMask = new Map<number, readonly GoogleFontsVariant[]>();
let familyMasks: Map<string, number> | undefined;

/**
 * Supported faces at weights 100, 200, ..., 900, normal first, then italic,
 * ordered by weight within each style. Matching is case-insensitive and
 * whitespace-normalised, as with `findGoogleFontsFamily`.
 *
 * Unknown families return an empty array. Weight-named aliases and metric clones
 * are not resolved. Results are immutable and shared between equivalent families.
 */
export function getGoogleFontsVariants(family: string): readonly GoogleFontsVariant[] {
	const canonical = findGoogleFontsFamily(family);
	if (!canonical) return EMPTY_VARIANTS;
	if (!familyMasks) {
		familyMasks = new Map(
			GOOGLE_FONTS_FAMILIES.map((name, index) => [
				name,
				parseInt(GOOGLE_FONTS_VARIANT_MASKS.slice(index * 4, index * 4 + 4), 36),
			]),
		);
	}
	const mask = familyMasks.get(canonical) ?? 0;
	const cached = variantsByMask.get(mask);
	if (cached) return cached;
	const variants: GoogleFontsVariant[] = [];
	for (const italic of [false, true]) {
		for (let weight = 100; weight <= 900; weight += 100) {
			if (mask & (1 << (weight / 100 - 1 + (italic ? 9 : 0)))) {
				variants.push(Object.freeze({ weight, italic }));
			}
		}
	}
	const result = Object.freeze(variants);
	variantsByMask.set(mask, result);
	return result;
}
