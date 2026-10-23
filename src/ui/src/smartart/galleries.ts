/**
 * The SmartArt Design galleries (Change Colors, SmartArt Styles) as `office-ui-gallery` state, for
 * any product: the catalogue comes from `ooxml-core/diagram`, the tiles from `./tiles`. A product
 * passes its theme's accent colours, the diagram's current scheme or style, and a translator
 * that receives the English name. PowerPoint builds its ribbon galleries from the same catalogue
 * and tiles (`pptx/render/ribbon-galleries/smartart-*-gallery.ts`).
 */
import {
	DIAGRAM_COLOR_ENTRIES,
	DIAGRAM_COLOR_SECTIONS,
	DIAGRAM_STYLE_ENTRIES,
	type DiagramColorEntry,
	type DiagramColorScheme,
	type DiagramStyleIntensity,
} from 'ooxml-core/diagram';
import type { OfficeGalleryState } from '../ribbon/gallery';
import { smartArtTileSvg, type SmartArtTileSize } from './tiles';

/** Resolved theme accents, `accent1` to `accent6`, as CSS colours. */
export type SmartArtAccents = Partial<Record<DiagramColorEntry['accents'][number], string>>;

/** Receives an English name and returns it translated. */
export type SmartArtGalleryTranslate = (english: string) => string;

const TILE: SmartArtTileSize = { width: 56, height: 36 };
const same: SmartArtGalleryTranslate = (english) => english;

/** Change Colors: one section per group of schemes, each tile filled with the scheme's accents. */
export function smartArtColorsGalleryState(
	accents: SmartArtAccents,
	current: DiagramColorScheme | undefined,
	t: SmartArtGalleryTranslate = same,
): OfficeGalleryState {
	return {
		id: 'smartArtColors',
		label: t('Change Colors'),
		sections: DIAGRAM_COLOR_SECTIONS.map((section) => ({
			title: t(section.title),
			columns: 5,
			tileWidth: TILE.width,
			tileHeight: TILE.height,
			items: DIAGRAM_COLOR_ENTRIES.filter((entry) => entry.section === section.id).map((entry) => ({
				id: entry.id,
				label: t(entry.name),
				applied: entry.id === current,
				preview: smartArtTileSvg(
					entry.id,
					entry.accents.map((accent) => accents[accent] ?? '#4472C4'),
					'flat',
					TILE,
				),
			})),
		})),
	};
}

/** SmartArt Styles: the style intensities the model carries, drawn in Accent 1. */
export function smartArtStylesGalleryState(
	accent1: string | undefined,
	current: DiagramStyleIntensity | undefined,
	t: SmartArtGalleryTranslate = same,
): OfficeGalleryState {
	return {
		id: 'smartArtStyles',
		label: t('SmartArt Styles'),
		sections: [
			{
				title: t('Best Match for Document'),
				columns: 5,
				tileWidth: TILE.width,
				tileHeight: TILE.height,
				items: DIAGRAM_STYLE_ENTRIES.map((entry) => ({
					id: entry.id,
					label: t(entry.name),
					applied: entry.id === current,
					preview: smartArtTileSvg(`style-${entry.id}`, [accent1 ?? '#4472C4'], entry.id, TILE),
				})),
			},
		],
	};
}
