import { describe, expect, it } from 'vitest';
import { resolveGoogleFontsFace } from 'ooxml-core/text';
import { buildFontCatalog } from './font-catalog';
import { fontFamilyItems } from './ribbon-home-state-extras';

const variants = (family: string) =>
	buildFontCatalog({ allFonts: [family], includeWeightVariants: true })[0]?.entries[0]?.variants;

describe('PPTX catalogue weight variants', () => {
	it('keeps the existing output unchanged unless requested', () => {
		expect(buildFontCatalog({ allFonts: ['DM Sans'] })[0]?.entries).toEqual([
			{ family: 'DM Sans' },
		]);
	});

	it('offers supported upright weights with the plain name for regular', () => {
		expect(variants('DM Sans')).toEqual([
			{ family: 'DM Sans Thin', weight: 100 },
			{ family: 'DM Sans ExtraLight', weight: 200 },
			{ family: 'DM Sans Light', weight: 300 },
			{ family: 'DM Sans', weight: 400 },
			{ family: 'DM Sans Medium', weight: 500 },
			{ family: 'DM Sans SemiBold', weight: 600 },
			{ family: 'DM Sans Bold', weight: 700 },
			{ family: 'DM Sans ExtraBold', weight: 800 },
			{ family: 'DM Sans Black', weight: 900 },
		]);
	});

	it('does not invent weights for sparse, single-weight or unknown families', () => {
		expect(variants('ADLaM Display')).toEqual([{ family: 'ADLaM Display', weight: 400 }]);
		expect(variants('ABeeZee')).toEqual([{ family: 'ABeeZee', weight: 400 }]);
		expect(variants('Arial')).toBeUndefined();
		expect(variants('Private Font')).toBeUndefined();
		expect(variants('DM Sans Medium')).toBeUndefined();
		expect(variants('Bungee Spice')).toEqual([{ family: 'Bungee Spice', weight: 400 }]);
	});

	it('normalises lookup but preserves the selected plain family spelling', () => {
		expect(variants('dm sans')).toContainEqual({ family: 'dm sans', weight: 400 });
		expect(variants('dm sans')).toContainEqual({ family: 'DM Sans Medium', weight: 500 });
	});

	it('never steals a separate catalogue family name for a weight', () => {
		expect(variants('Archivo')?.some((variant) => variant.family === 'Archivo Black')).toBe(false);
		for (const variant of variants('Archivo') ?? []) {
			expect(resolveGoogleFontsFace(variant.family)).toEqual(
				variant.weight === 400
					? { family: 'Archivo' }
					: { family: 'Archivo', weight: variant.weight },
			);
		}
	});

	it('does not infer local custom or embedded faces even when promoted to theme', () => {
		for (const local of [{ embeddedFonts: ['DM Sans'] }, { customFonts: [' dm sans '] }]) {
			const groups = buildFontCatalog({
				themeFonts: { body: 'DM Sans' },
				allFonts: ['DM Sans'],
				includeWeightVariants: true,
				...local,
			});
			expect(groups[0]?.entries).toEqual([{ family: 'DM Sans', themeRole: 'body' }]);
		}
	});

	it('shares the rows across bindings without duplicate or promoted family entries', () => {
		const rows = fontFamilyItems(
			{
				themeFonts: { heading: 'DM Sans' },
				customFonts: ['DM Sans Medium'],
				allFonts: ['dm sans', 'DM Sans Medium', 'ADLaM Display'],
			},
			(key) => key,
		);
		expect(rows.filter((row) => row.value.toLowerCase() === 'dm sans')).toHaveLength(1);
		expect(rows.filter((row) => row.value === 'DM Sans Medium')).toHaveLength(1);
		expect(rows.find((row) => row.value === 'DM Sans Medium')?.group).toBe(
			'pptx.font.group.custom',
		);
		expect(rows.find((row) => row.value === 'DM Sans Light')?.group).toBe('pptx.font.group.theme');
		expect(rows.some((row) => row.value === 'DM Sans Regular')).toBe(false);
		expect(rows.some((row) => row.value === 'ADLaM Display Medium')).toBe(false);
	});
});
