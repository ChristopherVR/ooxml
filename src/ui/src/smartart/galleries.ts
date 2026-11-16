/**
 * The SmartArt Design galleries (Change Colors, SmartArt Styles) as `office-ui-gallery` state, for
 * any product: Office's full built-in catalogue from `ooxml-core/diagram`
 * (`DIAGRAM_NATIVE_COLOR_ENTRIES`, `DIAGRAM_NATIVE_STYLE_ENTRIES`), the tiles from
 * `./definition-tiles`. Item ids are the definitions' short ids (`accent1_2`, `3d1`); a product
 * applies a pick with `builtinSmartArtFormatChange` of `ooxml-core/diagram/builtin`. A product
 * passes its theme colours, the diagram's applied colour or style id (a `uniqueId`, or an id
 * this package wrote for a model scheme), and a translator that receives the English name.
 */
import {
	DIAGRAM_NATIVE_COLOR_ENTRIES,
	DIAGRAM_NATIVE_COLOR_SECTIONS,
	DIAGRAM_NATIVE_STYLE_ENTRIES,
	DIAGRAM_NATIVE_STYLE_SECTIONS,
	appliedDiagramNativeId,
	type DiagramColorEntry,
} from 'ooxml-core/diagram';
import type { OfficeGalleryState } from '../ribbon/gallery';
import {
	smartArtColorDefinitionTileSvg,
	smartArtStyleDefinitionTileSvg,
	type SmartArtThemeColors,
} from './definition-tiles';
import type { SmartArtTileSize } from './tiles';

/** Resolved theme accents, `accent1` to `accent6`, as CSS colours. */
export type SmartArtAccents = Partial<Record<DiagramColorEntry['accents'][number], string>>;

/** Receives an English name and returns it translated. */
export type SmartArtGalleryTranslate = (english: string) => string;

const TILE: SmartArtTileSize = { width: 56, height: 36 };
const same: SmartArtGalleryTranslate = (english) => english;

/** Change Colors: Office's sections (Primary Theme Colors, Colorful, Accent 1 to 6). */
export function smartArtColorsGalleryState(
	theme: SmartArtAccents | SmartArtThemeColors,
	current: string | undefined,
	t: SmartArtGalleryTranslate = same,
): OfficeGalleryState {
	const applied = appliedDiagramNativeId('colors', current);
	return {
		id: 'smartArtColors',
		label: t('Change Colors'),
		sections: DIAGRAM_NATIVE_COLOR_SECTIONS.map((section) => ({
			title: t(section.title),
			columns: 5,
			tileWidth: TILE.width,
			tileHeight: TILE.height,
			items: DIAGRAM_NATIVE_COLOR_ENTRIES.filter((entry) => entry.category === section.id).map(
				(entry) => ({
					id: entry.id,
					label: t(entry.name),
					applied: entry.id === applied,
					preview: smartArtColorDefinitionTileSvg(entry, theme, TILE),
				}),
			),
		})),
	};
}

/** SmartArt Styles: Best Match for Document and 3-D, drawn in Accent 1. */
export function smartArtStylesGalleryState(
	accent1: string | undefined,
	current: string | undefined,
	t: SmartArtGalleryTranslate = same,
): OfficeGalleryState {
	const applied = appliedDiagramNativeId('style', current);
	return {
		id: 'smartArtStyles',
		label: t('SmartArt Styles'),
		sections: DIAGRAM_NATIVE_STYLE_SECTIONS.map((section) => ({
			title: t(section.title),
			columns: 5,
			tileWidth: TILE.width,
			tileHeight: TILE.height,
			items: DIAGRAM_NATIVE_STYLE_ENTRIES.filter((entry) => entry.category === section.id).map(
				(entry) => ({
					id: entry.id,
					label: t(entry.name),
					applied: entry.id === applied,
					preview: smartArtStyleDefinitionTileSvg(entry, accent1, TILE),
				}),
			),
		})),
	};
}
