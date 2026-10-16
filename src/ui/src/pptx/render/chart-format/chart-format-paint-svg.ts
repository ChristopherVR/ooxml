/**
 * SVG paint servers for the chart element fills and lines the flat chart
 * fields cannot carry (gradient, pattern and picture), as markup for the
 * self-contained effect images of `chart-format-effects-svg.ts`.
 *
 * Gradients come from the shared chart gradient painter
 * (`buildChartGradientDef`), patterns from the shared preset tiles
 * (`getPatternTile`), so a chart element paints exactly like a chart series
 * or a shape with the same fill.
 *
 * @module render/chart-format/chart-format-paint-svg
 */
import type { PptxChartGradientFill, PptxChartPatternFill } from 'ooxml-core/pptx';
import { buildChartGradientDef } from 'ooxml-core/chart';

import { getPatternTile } from '../fill-style';

/** A box in chart SVG coordinates. */
export interface PaintBox {
	x: number;
	y: number;
	w: number;
	h: number;
}

/** A paint: `defs` to put in the image and the `fill` / `stroke` value that uses them. */
export interface PaintServer {
	defs: string;
	paint: string;
}

const r4 = (v: number) => Math.round(v * 10000) / 10000;
const attr = (value: string) => value.replace(/[<>&"]/gu, '');

let serial = 0;
/** Restart paint ids (each effect image is its own document). */
export function resetPaintIds(): void {
	serial = 0;
}
const nextId = (prefix: string) => `${prefix}${++serial}`;

function stopsMarkup(
	stops: ReadonlyArray<{ offset: number; color: string; opacity?: number }>,
): string {
	return stops
		.map(
			(stop) =>
				`<stop offset="${r4(stop.offset)}" stop-color="${attr(stop.color)}"${stop.opacity !== undefined && stop.opacity < 1 ? ` stop-opacity="${r4(stop.opacity)}"` : ''}/>`,
		)
		.join('');
}

/**
 * A gradient paint server over `box`. Linear gradients are laid out in user
 * space on the box, so a horizontal or vertical line (a box with no height
 * or width) still shows its colours.
 */
export function gradientPaint(gradient: PptxChartGradientFill, box: PaintBox): PaintServer {
	const id = nextId('pg');
	const def = buildChartGradientDef(id, gradient, { width: box.w, height: box.h });
	if (def.kind === 'linearGradient') {
		const x = (u: number) => r4(box.x + u * box.w);
		const y = (v: number) => r4(box.y + v * box.h);
		return {
			defs: `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${x(def.x1)}" y1="${y(def.y1)}" x2="${x(def.x2)}" y2="${y(def.y2)}">${stopsMarkup(def.stops)}</linearGradient>`,
			paint: `url(#${id})`,
		};
	}
	if (def.kind === 'radialGradient') {
		return {
			defs: `<radialGradient id="${id}" cx="${r4(def.cx)}" cy="${r4(def.cy)}" r="${r4(def.r)}"${def.gradientTransform ? ` gradientTransform="${def.gradientTransform}"` : ''}>${stopsMarkup(def.stops)}</radialGradient>`,
			paint: `url(#${id})`,
		};
	}
	return {
		defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${r4(box.x)}" y="${r4(box.y)}" width="${r4(box.w)}" height="${r4(box.h)}"><image href="${attr(def.href)}" width="${r4(box.w)}" height="${r4(box.h)}" preserveAspectRatio="none"/></pattern>`,
		paint: `url(#${id})`,
	};
}

/** A preset pattern (`a:pattFill`) paint server, or `undefined` for an unknown preset. */
export function patternPaint(pattern: PptxChartPatternFill): PaintServer | undefined {
	const tile = getPatternTile(pattern.preset, pattern.foreground, pattern.background);
	if (!tile) {
		return undefined;
	}
	const id = nextId('pp');
	return {
		defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" width="${tile.w}" height="${tile.h}">${tile.inner}</pattern>`,
		paint: `url(#${id})`,
	};
}

/** Tile size for a tiled picture: the image's own size is not known synchronously. */
const PICTURE_TILE_PX = 48;

/** A picture paint server: stretched over `box`, or tiled. */
export function picturePaint(src: string, box: PaintBox, tile: boolean): PaintServer {
	const id = nextId('pi');
	const w = tile ? PICTURE_TILE_PX : box.w;
	const h = tile ? PICTURE_TILE_PX : box.h;
	return {
		defs: `<pattern id="${id}" patternUnits="userSpaceOnUse" x="${r4(box.x)}" y="${r4(box.y)}" width="${r4(w)}" height="${r4(h)}"><image href="${attr(src)}" width="${r4(w)}" height="${r4(h)}" preserveAspectRatio="${tile ? 'xMidYMid slice' : 'none'}"/></pattern>`,
		paint: `url(#${id})`,
	};
}
