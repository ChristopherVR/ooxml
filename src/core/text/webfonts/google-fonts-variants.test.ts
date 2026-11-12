import { describe, expect, it } from 'vitest';

import { GOOGLE_FONTS_FAMILIES, findGoogleFontsFamily, getGoogleFontsVariants } from '../index';
import { GOOGLE_FONTS_VARIANT_MASKS } from './google-fonts-catalogue';

const normal = (...weights: number[]) => weights.map((weight) => ({ weight, italic: false }));

describe('getGoogleFontsVariants', () => {
	it('exposes the supported variable-family standard weights, normal before italic', () => {
		const weights = normal(100, 200, 300, 400, 500, 600, 700, 800, 900);
		expect(getGoogleFontsVariants('DM Sans')).toEqual([
			...weights,
			...weights.map((face) => ({ ...face, italic: true })),
		]);
	});

	it('keeps a bounded variable range instead of advertising all nine weights', () => {
		const weights = normal(400, 500, 600, 700, 800, 900);
		expect(getGoogleFontsVariants('Alegreya')).toEqual([
			...weights,
			...weights.map((face) => ({ ...face, italic: true })),
		]);
	});

	it('does not fill gaps in a static family', () => {
		const weights = normal(100, 300, 400, 700, 900);
		expect(getGoogleFontsVariants('Lato')).toEqual([
			...weights,
			...weights.map((face) => ({ ...face, italic: true })),
		]);
	});

	it('keeps singleton families and families named with weight words at their actual weight', () => {
		expect(getGoogleFontsVariants('Bungee Shade')).toEqual(normal(400));
		expect(getGoogleFontsVariants('Archivo Black')).toEqual(normal(400));
	});

	it('includes only the italic faces the family supports', () => {
		expect(getGoogleFontsVariants('ABeeZee')).toEqual([
			{ weight: 400, italic: false },
			{ weight: 400, italic: true },
		]);
		expect(getGoogleFontsVariants('Abel')).toEqual(normal(400));
	});

	it('uses the existing case-insensitive and whitespace-normalised family lookup', () => {
		expect(getGoogleFontsVariants('  dm\tSANS  ')).toBe(getGoogleFontsVariants('DM Sans'));
		expect(findGoogleFontsFamily('  dm\tSANS  ')).toBe('DM Sans');
		expect(GOOGLE_FONTS_FAMILIES).toContain('DM Sans');
	});

	it('does not guess variants for unknown names, weight aliases, CSS lists or metric clones', () => {
		for (const family of [
			'',
			'Private Font',
			'DM Sans Medium',
			'"DM Sans", sans-serif',
			'Calibri',
		]) {
			expect(getGoogleFontsVariants(family)).toEqual([]);
		}
	});

	it('protects cached arrays and face objects from caller mutations', () => {
		const variants = getGoogleFontsVariants('DM Sans');
		expect(Object.isFrozen(variants)).toBe(true);
		expect(variants.every(Object.isFrozen)).toBe(true);
		expect(Reflect.set(variants, '0', { weight: 999, italic: true })).toBe(false);
		expect(Reflect.set(variants[0] ?? {}, 'weight', 999)).toBe(false);
		expect(getGoogleFontsVariants('DM Sans')[0]).toEqual({ weight: 100, italic: false });
	});

	it('has one valid, nonempty standard-face mask per catalogue family', () => {
		expect(GOOGLE_FONTS_FAMILIES.length).toBeGreaterThan(1000);
		expect(GOOGLE_FONTS_VARIANT_MASKS.length).toBe(GOOGLE_FONTS_FAMILIES.length * 4);
		for (const [index, family] of GOOGLE_FONTS_FAMILIES.entries()) {
			const packed = GOOGLE_FONTS_VARIANT_MASKS.slice(index * 4, index * 4 + 4);
			expect(packed).toMatch(/^[0-9a-z]{4}$/u);
			const mask = parseInt(packed, 36);
			expect(mask).toBeGreaterThan(0);
			expect(mask).toBeLessThan(1 << 18);
			const variants = getGoogleFontsVariants(family);
			expect(variants.length).toBeGreaterThan(0);
			expect(new Set(variants.map((face) => `${face.weight}:${face.italic}`)).size).toBe(
				variants.length,
			);
			expect(
				variants.every(
					(face) => face.weight >= 100 && face.weight <= 900 && face.weight % 100 === 0,
				),
			).toBe(true);
		}
	});
});
