/**
 * Design > Variants > Effects: PowerPoint's built-in effect schemes (`THEME_EFFECT_SCHEMES`, the
 * theme's `a:fmtScheme`). Each tile draws three shapes in the deck's accent colour with fill,
 * outline and effect styles 1 to 3 of that scheme, resolved through the handler's own style
 * resolver, so the tile shows what shapes in the deck will look like. A pick returns the scheme
 * for the binding to write into the theme.
 *
 * @module render/ribbon-galleries/theme-effects-gallery
 */
import type { ResolvedStyleMatrix, ShapeStyle, XmlObject } from 'ooxml-core/pptx';
import { findThemeEffectSchemeByName, THEME_EFFECT_SCHEMES } from 'ooxml-core/pptx';

import type { RibbonGalleryModule } from './gallery-module';
import { paintDefs, safeColor, svgId, svgTile } from './gallery-preview-svg';
import { galleryColorScheme } from './gallery-theme';
import type { RibbonGalleryContext, RibbonGalleryItem } from './gallery-types';

const TILE = { width: 76, height: 42 };
const SCHEME_LABEL_KEY = 'pptx.gallery.themeEffects.scheme';
const CELLS = [0, 1, 2] as const;

function num(value: number): string {
	return String(Math.round(value * 100) / 100);
}

function styleXml(idx: number, accent: string): XmlObject {
	const colour = (extra?: XmlObject) => ({ 'a:schemeClr': { '@_val': accent, ...extra } });
	return {
		'a:lnRef': { '@_idx': String(idx), ...colour() },
		'a:fillRef': { '@_idx': String(idx), ...colour() },
		'a:effectRef': { '@_idx': String(idx), ...colour() },
	};
}

/** One of the three sample shapes, drawn from its resolved `ShapeStyle`. */
function sampleShape(id: string, style: ShapeStyle, x: number, size: number): string {
	const gradient = style.fillMode === 'gradient' && (style.fillGradientStops?.length ?? 0) > 1;
	const fillId = `${svgId(id)}-f`;
	const paint = paintDefs(
		style.fillMode === 'none'
			? undefined
			: gradient
				? {
						stops: style.fillGradientStops!.map((stop) => ({
							color: stop.color,
							position: stop.position,
						})),
						angle: style.fillGradientAngle ?? 90,
					}
				: { color: style.fillColor, opacity: style.fillOpacity },
		fillId,
	);
	const strokeWidth = Math.min(Math.max(style.strokeWidth ?? 0, 0), 4);
	const stroke =
		strokeWidth > 0 && style.strokeColor && style.strokeColor !== 'transparent'
			? ` stroke="${safeColor(style.strokeColor, '#000000')}" stroke-width="${num(Math.max(1, strokeWidth))}"`
			: '';
	let defs = paint.defs;
	let filter = '';
	if (style.shadowColor) {
		const shadowId = `${svgId(id)}-s`;
		const dx = Math.max(-3, Math.min(3, style.shadowOffsetX ?? 0));
		const dy = Math.max(-3, Math.min(3, style.shadowOffsetY ?? 2));
		const blur = Math.min(style.shadowBlur ?? 3, 4) / 2;
		defs += `<filter id="${shadowId}" x="-30%" y="-30%" width="160%" height="170%"><feDropShadow dx="${num(dx)}" dy="${num(dy)}" stdDeviation="${num(blur)}" flood-color="${safeColor(style.shadowColor, '#000000')}" flood-opacity="${num(style.shadowOpacity ?? 0.4)}"/></filter>`;
		filter = ` filter="url(#${shadowId})"`;
	}
	const fillOpacity =
		!gradient && style.fillOpacity !== undefined && style.fillOpacity < 1
			? ` fill-opacity="${num(style.fillOpacity)}"`
			: '';
	const rect = `<rect x="${num(x)}" y="${num(TILE.height / 2 - size / 2 - 2)}" width="${num(size)}" height="${num(size)}" rx="3" fill="${paint.fill}"${fillOpacity}${stroke}${filter}/>`;
	return `${defs ? `<defs>${defs}</defs>` : ''}${rect}`;
}

/** Three shapes in the deck's accent colour under scheme `schemeId` (any when it is `undefined`). */
export function effectSchemeTileSvg(
	schemeId: string | undefined,
	ctx: RibbonGalleryContext,
	accent = 'accent1',
): string {
	const size = 15;
	const gap = (TILE.width - size * 3) / 4;
	const accentHex = galleryColorScheme(ctx)[accent as 'accent1'] ?? '#4472C4';
	const body = CELLS.map((index) => {
		const resolved: ResolvedStyleMatrix | undefined = ctx.resolveStyleMatrix?.(
			styleXml(index + 1, accent),
			schemeId,
		);
		const style: ShapeStyle = resolved?.shapeStyle ?? {
			fillMode: 'solid',
			fillColor: accentHex,
			strokeColor: accentHex,
			strokeWidth: 1,
		};
		return sampleShape(`${schemeId ?? 'deck'}-${index}`, style, gap + index * (size + gap), size);
	}).join('');
	return svgTile(TILE.width, TILE.height, '', body);
}

function item(id: string, name: string, ctx: RibbonGalleryContext, applied: boolean) {
	return {
		id,
		labelKey: SCHEME_LABEL_KEY,
		labelParams: { name },
		label: `Effects: ${name}`,
		previewSvg: effectSchemeTileSvg(id, ctx),
		applied,
	} satisfies RibbonGalleryItem;
}

export const THEME_EFFECTS_GALLERY: RibbonGalleryModule = {
	build(ctx) {
		const active = findThemeEffectSchemeByName(ctx.theme?.formatScheme?.name);
		return {
			id: 'themeEffects',
			labelKey: 'pptx.gallery.themeEffects.title',
			label: 'Effects',
			disabled: false,
			previewable: true,
			sections: [
				{
					id: 'builtIn',
					titleKey: 'pptx.gallery.themeEffects.section.builtIn',
					title: 'Built-in',
					columns: 3,
					tileWidth: TILE.width,
					tileHeight: TILE.height,
					items: THEME_EFFECT_SCHEMES.map((scheme) =>
						item(scheme.id, scheme.name, ctx, scheme.id === active?.id),
					),
				},
			],
		};
	},
	apply(itemId) {
		const scheme = THEME_EFFECT_SCHEMES.find((candidate) => candidate.id === itemId);
		return scheme ? { kind: 'themeEffects', schemeId: scheme.id, name: scheme.name } : null;
	},
};
