import {
	findGoogleFontsFamily,
	getGoogleFontsVariants,
	resolveGoogleFontsFace,
} from 'ooxml-core/text';
import { buildFontCatalog as buildSharedCatalog } from '../../fonts/catalog';
import type {
	FontCatalogInput as SharedFontCatalogInput,
	FontCatalogEntry as SharedFontCatalogEntry,
	FontCatalogGroup as SharedFontCatalogGroup,
} from '../../fonts/catalog';
import { COMMON_FONT_FAMILIES } from './text-format-presets';
export type { FontCatalogGroupId } from '../../fonts/catalog';

/** One supported upright Google Fonts weight, stored as a PowerPoint typeface. */
export interface FontCatalogVariant {
	/** The family name to write into the run's fontFamily, not a CSS stack. */
	family: string;
	/** The CSS weight selected by this name. */
	weight: number;
}

export interface FontCatalogEntry extends SharedFontCatalogEntry {
	/** Supported upright weights, including the plain family for 400, when requested. */
	variants?: readonly FontCatalogVariant[];
}

export interface FontCatalogGroup extends Omit<SharedFontCatalogGroup, 'entries'> {
	entries: FontCatalogEntry[];
}

export interface FontCatalogInput extends SharedFontCatalogInput {
	/** Include supported Google Fonts weights. Defaults to false for compatibility. */
	includeWeightVariants?: boolean;
}

const WEIGHT_NAMES: Readonly<Record<number, string>> = {
	100: 'Thin',
	200: 'ExtraLight',
	300: 'Light',
	500: 'Medium',
	600: 'SemiBold',
	700: 'Bold',
	800: 'ExtraBold',
	900: 'Black',
};

/**
 * Group the PPTX picker fonts. Opt in to variants to receive only weights the
 * bundled Google Fonts metadata supports. Local embedded/custom families and
 * unknown names keep their own faces; their availability cannot be inferred.
 * This runtime entry is DOM-free: ooxml-ui/pptx/render/font-catalog.
 */
export function buildFontCatalog(input: FontCatalogInput = {}): FontCatalogGroup[] {
	const groups = buildSharedCatalog({
		...input,
		allFonts: input.allFonts ?? COMMON_FONT_FAMILIES,
		labelKeys: {
			theme: 'pptx.font.group.theme',
			embedded: 'pptx.font.group.embedded',
			custom: 'pptx.font.group.custom',
			all: 'pptx.font.group.all',
		},
	});
	if (!input.includeWeightVariants) return groups;
	const localFamilies = new Set(
		[...(input.embeddedFonts ?? []), ...(input.customFonts ?? [])].map((name) =>
			name.trim().replace(/\s+/gu, ' ').toLowerCase(),
		),
	);
	return groups.map((group) => ({
		...group,
		entries: group.entries.map((entry) => {
			const canonical = findGoogleFontsFamily(entry.family);
			if (!canonical || localFamilies.has(canonical.toLowerCase())) return entry;
			const variants = getGoogleFontsVariants(canonical)
				.filter((face) => !face.italic)
				.flatMap(({ weight }) => {
					const family = weight === 400 ? entry.family : `${canonical} ${WEIGHT_NAMES[weight]}`;
					// A catalogue family such as Archivo Black owns its name. Do not
					// offer it as another family's weight when it resolves elsewhere.
					const face = resolveGoogleFontsFace(family);
					return face?.family === canonical && (face.weight ?? 400) === weight
						? [{ family, weight }]
						: [];
				});
			return variants.length ? { ...entry, variants } : entry;
		}),
	}));
}

export function resolveDefaultFontFamily(
	placeholderType: string | undefined,
	themeFonts: { heading?: string; body?: string } | undefined,
	fallback = 'Segoe UI',
): string {
	const normalized = placeholderType?.trim().toLowerCase();
	const isHeading = normalized === 'title' || normalized === 'ctrtitle';
	const preferred = isHeading ? themeFonts?.heading : themeFonts?.body;
	return preferred?.trim() || themeFonts?.body?.trim() || themeFonts?.heading?.trim() || fallback;
}
