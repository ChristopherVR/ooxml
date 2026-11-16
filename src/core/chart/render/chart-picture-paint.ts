/**
 * Picture fills (`a:blipFill`) of chart elements as SVG paint: the series, a point, the plot area
 * and the chart area. A picture is stretched over the element's box (`a:stretch`, with its fill
 * rectangle), tiled at its own size scaled by `a:tile` (`sx`, `sy`, offset, alignment), or, for a
 * bar or column whose `c:pictureOptions` asks for it, stacked: repeated along the bar at the bar's
 * width (`stack`) or once per `c:pictureStackUnit` value-axis units (`stackScale`). Tiles mirror
 * by `a:tile/@flip`; `a:srcRect` crops (or, negative, pads) the picture before it is placed, and
 * `a:alphaModFix` sets its opacity. The picture's bytes come from the host, which resolves the
 * fill's relationship id to a URL (`ChartSpace.pictureSources` for a chart part).
 *
 * Not painted: picture effects other than transparency.
 *
 * @module chart/render/chart-picture-paint
 */
import type { DrawingFill } from '../../drawingml/types';
import { EMU_PER_PIXEL } from '../../units/constants';
import { esc, n } from './chart-svg-util';
import { chartPictureSize } from './chart-picture-size';

export { chartPictureSize } from './chart-picture-size';

/** How a picture covers its element. */
export type ChartPictureMode = 'stretch' | 'tile' | 'stack';

/** A picture fill resolved for painting. */
export interface ChartPicturePaint {
	href: string;
	mode: ChartPictureMode;
	/** The picture's own size in CSS pixels (96 dpi), when its header could be read. */
	width: number;
	height: number;
	opacity?: number;
	/** Tile scale (fractions), offset (pixels), alignment and mirroring. */
	tile?: {
		scaleX: number;
		scaleY: number;
		offsetX: number;
		offsetY: number;
		alignment: string;
		flip: 'none' | 'x' | 'y' | 'xy';
	};
	/** Stretch fill rectangle insets, fractions of the box. */
	inset?: Insets;
	/** `a:srcRect`: the part of the picture shown, as insets of the picture (negative pads it). */
	crop?: Insets;
	/** `stackScale`: the value-axis units one stacked picture stands for. */
	stackUnit?: number;
}

interface Insets {
	l: number;
	t: number;
	r: number;
	b: number;
}

/** Resolves a picture fill's relationship id to a URL (a `data:` URL for a package part). */
export type ChartPictureSource = (relId: string) => string | undefined;

const DEFAULT_SIZE = 96;

/** Whether `c:pictureOptions` (as written) stacks the picture along the bar. */
export function chartPictureOptionsStack(xml: string | undefined): boolean {
	return /pictureFormat\b[^>]*\bval\s*=\s*["']stack(?:Scale)?["']/u.test(xml ?? '');
}

/** The `c:pictureStackUnit` of a `stackScale` picture format, when positive. */
export function chartPictureOptionsUnit(xml: string | undefined): number | undefined {
	if (!/pictureFormat\b[^>]*\bval\s*=\s*["']stackScale["']/u.test(xml ?? '')) return undefined;
	const unit = Number(/pictureStackUnit\b[^>]*\bval\s*=\s*["']([^"']+)["']/u.exec(xml ?? '')?.[1]);
	return Number.isFinite(unit) && unit > 0 ? unit : undefined;
}

/** The paint of a picture fill, or undefined for another fill or a picture the host lacks. */
export function chartPicturePaint(
	fill: DrawingFill | undefined,
	source: ChartPictureSource | undefined,
	stack = false,
	stackUnit?: number,
): ChartPicturePaint | undefined {
	if (fill?.kind !== 'picture' || !fill.relId || !source) return undefined;
	const href = source(fill.relId);
	if (!href) return undefined;
	const size = chartPictureSize(href);
	const crop = fill.crop && {
		l: fill.crop.cropLeft ?? 0,
		t: fill.crop.cropTop ?? 0,
		r: fill.crop.cropRight ?? 0,
		b: fill.crop.cropBottom ?? 0,
	};
	const shown = crop && 1 - crop.l - crop.r > 0 && 1 - crop.t - crop.b > 0 ? crop : undefined;
	const tile = fill.layout?.tile ?? (/<a:tile\b/u.test(fill.sourceXml ?? '') ? {} : undefined);
	const paint: ChartPicturePaint = {
		href,
		mode: tile ? 'tile' : stack ? 'stack' : 'stretch',
		width: (size?.width || DEFAULT_SIZE) * (shown ? 1 - shown.l - shown.r : 1),
		height: (size?.height || DEFAULT_SIZE) * (shown ? 1 - shown.t - shown.b : 1),
	};
	if (shown && (shown.l || shown.t || shown.r || shown.b)) paint.crop = shown;
	if (!tile && stack && stackUnit) paint.stackUnit = stackUnit;
	const alpha = fill.effects?.find((effect) => effect.kind === 'alphaModFix');
	if (alpha?.kind === 'alphaModFix' && alpha.amount < 1) paint.opacity = Math.max(0, alpha.amount);
	if (tile)
		paint.tile = {
			scaleX: tile.scaleX ?? 1,
			scaleY: tile.scaleY ?? 1,
			offsetX: (tile.offsetXEmu ?? 0) / EMU_PER_PIXEL,
			offsetY: (tile.offsetYEmu ?? 0) / EMU_PER_PIXEL,
			alignment: tile.alignment ?? 'tl',
			flip: tile.flip ?? 'none',
		};
	const rect = fill.layout?.fillRect;
	if (!tile && rect && (rect.l || rect.t || rect.r || rect.b)) paint.inset = { ...rect };
	return paint;
}

/** A box in user space, for tiles anchored to an element and for stacked bars. */
export interface ChartPictureBox {
	x: number;
	y: number;
	w: number;
	h: number;
	/** Stacking runs from this edge (the bar's base). */
	from?: 'bottom' | 'top' | 'left' | 'right';
	/** `stackScale`: the length along the bar of one picture, in user units. */
	unitLength?: number;
}

/** The picture in the box `x`, `y`, `w`, `h`: cropped to its `a:srcRect` through a viewport. */
function image(paint: ChartPicturePaint, x: number, y: number, w: number, h: number): string {
	const opacity = paint.opacity === undefined ? '' : ` opacity="${n(paint.opacity)}"`;
	const crop = paint.crop;
	if (!crop)
		return `<image href="${esc(paint.href)}" x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" preserveAspectRatio="none"${opacity}/>`;
	const view = [crop.l, crop.t, 1 - crop.l - crop.r, 1 - crop.t - crop.b].map(n).join(' ');
	return `<svg x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" viewBox="${view}" preserveAspectRatio="none" overflow="hidden"><image href="${esc(paint.href)}" width="1" height="1" preserveAspectRatio="none"${opacity}/></svg>`;
}

/** One tile, and its mirrored copies when the tile flips (`x`, `y` or `xy`). */
function tileImages(paint: ChartPicturePaint, w: number, h: number, flip: string): string {
	const one = image(paint, 0, 0, w, h);
	const mirrored = (a: number, d: number, e: number, f: number) =>
		`<g transform="matrix(${a} 0 0 ${d} ${n(e)} ${n(f)})">${one}</g>`;
	const x = flip === 'x' || flip === 'xy';
	const y = flip === 'y' || flip === 'xy';
	return (
		one +
		(x ? mirrored(-1, 1, 2 * w, 0) : '') +
		(y ? mirrored(1, -1, 0, 2 * h) : '') +
		(x && y ? mirrored(-1, -1, 2 * w, 2 * h) : '')
	);
}

function alignOffset(alignment: string, box: ChartPictureBox, w: number, h: number) {
	const centreX = alignment === 't' || alignment === 'b' || alignment === 'ctr';
	const centreY = alignment === 'l' || alignment === 'r' || alignment === 'ctr';
	return {
		x: centreX ? (box.w - w) / 2 : alignment.endsWith('r') ? box.w - w : 0,
		y: centreY ? (box.h - h) / 2 : alignment.startsWith('b') ? box.h - h : 0,
	};
}

/**
 * The `<pattern>` element `id` paints with. Without `box` a stretched picture covers each element's
 * own bounding box and a tile starts at the user-space origin; with `box` the pattern is laid out
 * for that one element (a tile aligned in it, a stack along it).
 */
export function chartPicturePattern(
	id: string,
	paint: ChartPicturePaint,
	box?: ChartPictureBox,
): string {
	if (paint.mode === 'stack' && box && box.w > 0 && box.h > 0) {
		const vertical = box.from !== 'left' && box.from !== 'right';
		const across = vertical ? box.w : box.h;
		const along =
			box.unitLength && box.unitLength > 0
				? box.unitLength
				: vertical
					? (across * paint.height) / paint.width
					: (across * paint.width) / paint.height;
		const w = vertical ? across : along;
		const h = vertical ? along : across;
		const x = box.from === 'right' ? box.x + box.w : box.x;
		const y = box.from === 'bottom' || box.from === undefined ? box.y + box.h : box.y;
		return `<pattern id="${esc(id)}" patternUnits="userSpaceOnUse" x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}">${image(paint, 0, 0, w, h)}</pattern>`;
	}
	if (paint.mode === 'tile' && paint.tile) {
		const w = Math.max(1, paint.width * paint.tile.scaleX);
		const h = Math.max(1, paint.height * paint.tile.scaleY);
		const origin = box ? alignOffset(paint.tile.alignment, box, w, h) : { x: 0, y: 0 };
		const x = (box?.x ?? 0) + origin.x + paint.tile.offsetX;
		const y = (box?.y ?? 0) + origin.y + paint.tile.offsetY;
		const flip = paint.tile.flip;
		const pw = flip === 'x' || flip === 'xy' ? 2 * w : w;
		const ph = flip === 'y' || flip === 'xy' ? 2 * h : h;
		return `<pattern id="${esc(id)}" patternUnits="userSpaceOnUse" x="${n(x)}" y="${n(y)}" width="${n(pw)}" height="${n(ph)}">${tileImages(paint, w, h, flip)}</pattern>`;
	}
	const inset = paint.inset ?? { l: 0, t: 0, r: 0, b: 0 };
	return `<pattern id="${esc(id)}" patternUnits="objectBoundingBox" patternContentUnits="objectBoundingBox" width="1" height="1">${image(paint, inset.l, inset.t, 1 - inset.l - inset.r, 1 - inset.t - inset.b)}</pattern>`;
}
