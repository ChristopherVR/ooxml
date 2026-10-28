import { buildFontCatalog as buildSharedCatalog } from '../../fonts/catalog';
import type { FontCatalogInput } from '../../fonts/catalog';
import { COMMON_FONT_FAMILIES } from './text-format-presets';
export type {
	FontCatalogInput,
	FontCatalogEntry,
	FontCatalogGroup,
	FontCatalogGroupId,
} from '../../fonts/catalog';
export function buildFontCatalog(input: FontCatalogInput = {}) {
	return buildSharedCatalog({
		...input,
		allFonts: input.allFonts ?? COMMON_FONT_FAMILIES,
		labelKeys: {
			theme: 'pptx.font.group.theme',
			embedded: 'pptx.font.group.embedded',
			custom: 'pptx.font.group.custom',
			all: 'pptx.font.group.all',
		},
	});
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
