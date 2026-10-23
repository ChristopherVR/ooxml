/**
 * Design > Themes: the inline Themes gallery. Tiles are the built-in themes the Browse Themes menu
 * already offers (`GALLERY_THEME_PRESETS`), drawn as a small slide in the theme's own colours and
 * heading font; a pick returns the whole theme (colour and font scheme) for the binding's
 * "apply theme" path, the same one the Browse Themes menu runs.
 *
 * @module render/ribbon-galleries/themes-gallery
 */
import type { PptxThemeColorScheme, PptxThemeFontScheme, PptxThemePreset } from 'ooxml-core/pptx';

import { activeGalleryThemePreset, GALLERY_THEME_PRESETS } from '../../theme/theme-gallery-presets';
import type { RibbonGalleryModule } from './gallery-module';
import { escapeSvgText, safeColor, svgTile } from './gallery-preview-svg';
import { galleryColorScheme } from './gallery-theme';
import type { RibbonGalleryContext, RibbonGalleryItem } from './gallery-types';

const TILE = { width: 88, height: 50 };
const CURRENT_ID = 'current';
const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

/** A slide in miniature: background, "Aa" in the heading font and the six accents. */
export function themeTileSvg(
	scheme: PptxThemeColorScheme,
	headingFont: string | undefined,
	size: { width: number; height: number } = TILE,
): string {
	const face = escapeSvgText((headingFont ?? 'Calibri').replace(/[;{}<>]/gu, ''));
	const swatchWidth = TILE.width / 6;
	const swatches = ACCENTS.map(
		(key, index) =>
			`<rect x="${(index * swatchWidth).toFixed(2)}" y="${TILE.height - 8}" width="${swatchWidth.toFixed(2)}" height="8" fill="${safeColor(scheme[key], '#808080')}"/>`,
	).join('');
	const body =
		`<rect x="0.5" y="0.5" width="${TILE.width - 1}" height="${TILE.height - 1}" fill="${safeColor(scheme.lt1, '#FFFFFF')}" stroke="#C8C8C8"/>` +
		`<rect x="0" y="0" width="${TILE.width}" height="5" fill="${safeColor(scheme.accent1, '#808080')}"/>` +
		`<text x="8" y="31" font-family="${face}, Calibri, Arial, sans-serif" font-size="22" fill="${safeColor(scheme.dk2, '#000000')}">Aa</text>` +
		`<rect x="48" y="20" width="32" height="3" fill="${safeColor(scheme.dk1, '#000000')}" opacity="0.55"/>` +
		`<rect x="48" y="27" width="24" height="3" fill="${safeColor(scheme.dk1, '#000000')}" opacity="0.35"/>` +
		swatches;
	// Drawn on the Themes tile grid and scaled to `size`, so the smaller Variants tiles match.
	return svgTile(TILE.width, TILE.height, '', body).replace(
		`width="${TILE.width}" height="${TILE.height}"`,
		`width="${size.width}" height="${size.height}"`,
	);
}

function tile(preset: PptxThemePreset, applied: boolean): RibbonGalleryItem {
	return {
		id: preset.id,
		labelKey: 'pptx.gallery.themes.theme',
		labelParams: { name: preset.name },
		label: `Theme: ${preset.name}`,
		previewSvg: themeTileSvg(preset.colorScheme, preset.fontScheme.majorFont?.latin),
		applied,
	};
}

function currentFonts(ctx: RibbonGalleryContext): PptxThemeFontScheme | undefined {
	return ctx.theme?.fontScheme;
}

export const THEMES_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const active = activeGalleryThemePreset(ctx.theme);
		const items = GALLERY_THEME_PRESETS.map((preset) => tile(preset, preset.id === active?.id));
		if (ctx.theme) {
			// The deck's theme always heads the list, as PowerPoint's "This Presentation". Showing it
			// only for a non-built-in theme shifted every tile when a built-in was picked, so the tile
			// under the pointer became the next theme and started previewing it.
			const scheme = galleryColorScheme(ctx);
			items.unshift({
				id: CURRENT_ID,
				labelKey: 'pptx.gallery.themes.current',
				label: 'Current theme',
				previewSvg: themeTileSvg(scheme, currentFonts(ctx)?.majorFont?.latin),
				applied: !active,
			});
		}
		return {
			id: 'themes',
			labelKey: 'pptx.gallery.themes.title',
			label: 'Themes',
			disabled: false,
			previewable: true,
			inlineCount: 3,
			sections: [
				{
					id: 'builtIn',
					titleKey: 'pptx.gallery.themes.section.builtIn',
					title: 'Built-in',
					columns: 4,
					tileWidth: TILE.width,
					tileHeight: TILE.height,
					items,
				},
			],
		};
	},
	apply(itemId) {
		// Applying the theme the deck already has is harmless and resets edited fonts.
		return fromPreset(GALLERY_THEME_PRESETS.find((candidate) => candidate.id === itemId));
	},
};

function fromPreset(preset: PptxThemePreset | undefined) {
	return preset
		? ({
				kind: 'theme',
				colorScheme: { ...preset.colorScheme },
				fontScheme: structuredClone(preset.fontScheme),
				name: preset.name,
			} as const)
		: null;
}
