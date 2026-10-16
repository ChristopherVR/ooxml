/**
 * Design > Variants > Background Styles: PowerPoint's 4 by 3 grid. Each row is one entry of the
 * theme's background fill styles (`bgRef idx` 1001 to 1003), each column a colour (light 1, light
 * 2, then the dark pair). A tile draws the style as the deck's theme resolves it; a pick applies
 * it to every slide through the master, the context menu offers "Apply to Selected Slides", and
 * "Format Background..." opens the background pane.
 *
 * @module render/ribbon-galleries/theme-backgrounds-gallery
 */
import type { ShapeStyle, XmlObject } from 'ooxml-core/pptx';
import { BACKGROUND_STYLE_PRESETS } from 'ooxml-core/pptx';
import type { PptxBackgroundStylePreset } from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import { safeColor, svgId, svgTile } from './gallery-preview-svg';
import { galleryColorScheme } from './gallery-theme';
import type { RibbonGalleryContext, RibbonGalleryItem } from './gallery-types';

const TILE = { width: 52, height: 36 };
const LABEL_KEY = 'pptx.gallery.themeBackgrounds.style';

function num(value: number): string {
	return String(Math.round(value * 100) / 100);
}

/** The scheme slot a preset's colour is on screen: the dark columns swap `bg1`/`bg2` to `dk1`/`dk2`. */
function presetSlot(preset: PptxBackgroundStylePreset): 'lt1' | 'lt2' | 'dk1' | 'dk2' {
	if (preset.dark) {
		return preset.scheme === 'bg1' ? 'dk1' : 'dk2';
	}
	return preset.scheme === 'bg1' ? 'lt1' : 'lt2';
}

function bgRefXml(preset: PptxBackgroundStylePreset): XmlObject {
	return {
		'a:fillRef': {
			'@_idx': String(preset.idx),
			'a:schemeClr': { '@_val': presetSlot(preset) },
		},
	};
}

/** The fill SVG definitions and `fill` value for a resolved background style. */
function backgroundPaint(
	id: string,
	style: ShapeStyle | undefined,
	fallback: string,
): { defs: string; fill: string } {
	const stops = style?.fillGradientStops ?? [];
	if (style?.fillMode === 'gradient' && stops.length > 1) {
		const body = stops
			.map(
				(stop) =>
					`<stop offset="${num(stop.position)}%" stop-color="${safeColor(stop.color, fallback)}"/>`,
			)
			.join('');
		if (style.fillGradientType === 'radial') {
			return {
				defs: `<radialGradient id="${id}" cx="50%" cy="50%" r="75%">${body}</radialGradient>`,
				fill: `url(#${id})`,
			};
		}
		const rad = (((style.fillGradientAngle ?? 90) % 360) * Math.PI) / 180;
		const x = Math.cos(rad) / 2;
		const y = Math.sin(rad) / 2;
		return {
			defs: `<linearGradient id="${id}" x1="${num(0.5 - x)}" y1="${num(0.5 - y)}" x2="${num(0.5 + x)}" y2="${num(0.5 + y)}">${body}</linearGradient>`,
			fill: `url(#${id})`,
		};
	}
	return { defs: '', fill: safeColor(style?.fillColor, fallback) };
}

/** A slide in miniature filled with background style `preset` and a few title/body lines. */
export function backgroundStyleTileSvg(
	preset: PptxBackgroundStylePreset,
	ctx: RibbonGalleryContext,
): string {
	const scheme = galleryColorScheme(ctx);
	const slot = presetSlot(preset);
	const fallback = scheme[slot];
	const resolved = ctx.resolveStyleMatrix?.(bgRefXml(preset));
	const id = svgId(`bg-style-${preset.index}`);
	const paint = backgroundPaint(`${id}-f`, resolved?.shapeStyle, fallback);
	// Text on the dark pair is light, as the swapped colour map makes it.
	const text = safeColor(preset.dark ? scheme.lt1 : scheme.dk1, '#000000');
	const body =
		`<rect x="0.5" y="0.5" width="${TILE.width - 1}" height="${TILE.height - 1}" fill="${paint.fill}" stroke="#C8C8C8"/>` +
		`<rect x="6" y="9" width="26" height="4" fill="${text}" opacity="0.85"/>` +
		`<rect x="6" y="18" width="38" height="2.5" fill="${text}" opacity="0.5"/>` +
		`<rect x="6" y="24" width="30" height="2.5" fill="${text}" opacity="0.5"/>`;
	return svgTile(TILE.width, TILE.height, paint.defs, body);
}

/** The preset the active slide uses, when it picked one from this gallery. */
function activePreset(ctx: RibbonGalleryContext): PptxBackgroundStylePreset | undefined {
	const ref = ctx.activeSlide?.backgroundStyleRef;
	if (ref?.kind !== 'preset') {
		return undefined;
	}
	const dark = ctx.activeSlide?.clrMapOverride?.['bg1'] === 'dk1';
	return BACKGROUND_STYLE_PRESETS.find(
		(preset) => preset.idx === ref.idx && preset.scheme === ref.scheme && preset.dark === dark,
	);
}

function tile(
	preset: PptxBackgroundStylePreset,
	ctx: RibbonGalleryContext,
	active: PptxBackgroundStylePreset | undefined,
): RibbonGalleryItem {
	return {
		id: String(preset.index),
		labelKey: LABEL_KEY,
		labelParams: { number: preset.index },
		label: `Style ${preset.index}`,
		previewSvg: backgroundStyleTileSvg(preset, ctx),
		applied: active?.index === preset.index,
		actions: [
			{
				id: 'all',
				labelKey: 'pptx.gallery.themeBackgrounds.applyAll',
				label: 'Apply to All Slides',
			},
			{
				id: 'selected',
				labelKey: 'pptx.gallery.themeBackgrounds.applySelected',
				label: 'Apply to Selected Slides',
			},
		],
	};
}

export const THEME_BACKGROUNDS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const active = activePreset(ctx);
		return {
			id: 'themeBackgrounds',
			labelKey: 'pptx.gallery.themeBackgrounds.title',
			label: 'Background Styles',
			disabled: false,
			previewable: true,
			sections: [
				{
					id: 'styles',
					columns: 4,
					tileWidth: TILE.width,
					tileHeight: TILE.height,
					items: BACKGROUND_STYLE_PRESETS.map((preset) => tile(preset, ctx, active)),
				},
			],
			footer: [
				{
					id: 'formatBackground',
					labelKey: 'pptx.gallery.themeBackgrounds.formatBackground',
					label: 'Format Background...',
				},
			],
		};
	},
	apply(itemId, _ctx, action) {
		if (itemId === 'formatBackground') {
			return { kind: 'command', command: 'formatBackground' };
		}
		const index = Number.parseInt(itemId, 10);
		if (!BACKGROUND_STYLE_PRESETS.some((preset) => preset.index === index)) {
			return null;
		}
		return {
			kind: 'backgroundStyle',
			styleIndex: index,
			scope: action === 'selected' ? 'selected' : 'all',
		};
	},
};
