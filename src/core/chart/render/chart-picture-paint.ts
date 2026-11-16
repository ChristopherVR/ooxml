/**
 * Picture fills (`a:blipFill`) of chart elements as SVG paint: the series, a point, the plot area
 * and the chart area. A picture is stretched over the element's box (`a:stretch`, with its fill
 * rectangle), tiled at its own size scaled by `a:tile` (`sx`, `sy`, offset, alignment), or, for a
 * bar or column whose `c:pictureOptions` asks for it, stacked: repeated along the bar at the bar's
 * width. `a:alphaModFix` sets the picture's opacity. The picture's bytes come from the host, which
 * resolves the fill's relationship id to a URL (`ChartSpace.pictureSources` for a chart part).
 *
 * Not painted: tile mirroring (`flip`), `stackScale` units (stacked like `stack`), picture
 * effects other than transparency, and crop rectangles (`a:srcRect`).
 *
 * @module chart/render/chart-picture-paint
 */
import type { DrawingFill } from '../../drawingml/types';
import { EMU_PER_PIXEL } from '../../units/constants';
import { esc, n } from './chart-svg-util';

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
	/** Tile scale (fractions), offset (pixels) and alignment. */
	tile?: { scaleX: number; scaleY: number; offsetX: number; offsetY: number; alignment: string };
	/** Stretch fill rectangle insets, fractions of the box. */
	inset?: { l: number; t: number; r: number; b: number };
}

/** Resolves a picture fill's relationship id to a URL (a `data:` URL for a package part). */
export type ChartPictureSource = (relId: string) => string | undefined;

const DEFAULT_SIZE = 96;

function decodeBase64Head(url: string, bytes: number): Uint8Array | undefined {
	const match = /^data:[^;,]*;base64,/u.exec(url);
	if (!match) return undefined;
	const chars = Math.ceil((bytes * 4) / 3 / 4) * 4;
	try {
		const binary = atob(url.slice(match[0].length, match[0].length + chars).replace(/\s/gu, ''));
		return Uint8Array.from(binary, (char) => char.charCodeAt(0));
	} catch {
		return undefined;
	}
}

/** The pixel size a PNG, GIF, BMP or JPEG `data:` URL declares in its header, if readable. */
export function chartPictureSize(url: string): { width: number; height: number } | undefined {
	const head = decodeBase64Head(url, 64 * 1024);
	if (!head || head.length < 10) return undefined;
	const view = new DataView(head.buffer, head.byteOffset, head.byteLength);
	if (head[0] === 0x89 && head[1] === 0x50 && head[2] === 0x4e && head[3] === 0x47)
		return head.length < 24 ? undefined : { width: view.getUint32(16), height: view.getUint32(20) };
	if (head[0] === 0x47 && head[1] === 0x49 && head[2] === 0x46)
		return { width: view.getUint16(6, true), height: view.getUint16(8, true) };
	if (head[0] === 0x42 && head[1] === 0x4d && head.length >= 26)
		return { width: Math.abs(view.getInt32(18, true)), height: Math.abs(view.getInt32(22, true)) };
	if (head[0] === 0xff && head[1] === 0xd8) {
		let at = 2;
		while (at + 9 < head.length) {
			if (head[at] !== 0xff) return undefined;
			const marker = head[at + 1]!;
			const length = view.getUint16(at + 2);
			if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker))
				return { width: view.getUint16(at + 7), height: view.getUint16(at + 5) };
			at += 2 + length;
		}
	}
	return undefined;
}

/** Whether `c:pictureOptions` (as written) stacks the picture along the bar. */
export function chartPictureOptionsStack(xml: string | undefined): boolean {
	return /pictureFormat\b[^>]*\bval\s*=\s*["']stack(?:Scale)?["']/u.test(xml ?? '');
}

/** The paint of a picture fill, or undefined for another fill or a picture the host lacks. */
export function chartPicturePaint(
	fill: DrawingFill | undefined,
	source: ChartPictureSource | undefined,
	stack = false,
): ChartPicturePaint | undefined {
	if (fill?.kind !== 'picture' || !fill.relId || !source) return undefined;
	const href = source(fill.relId);
	if (!href) return undefined;
	const size = chartPictureSize(href);
	const tile = fill.layout?.tile ?? (/<a:tile\b/u.test(fill.sourceXml ?? '') ? {} : undefined);
	const paint: ChartPicturePaint = {
		href,
		mode: tile ? 'tile' : stack ? 'stack' : 'stretch',
		width: size?.width || DEFAULT_SIZE,
		height: size?.height || DEFAULT_SIZE,
	};
	const alpha = fill.effects?.find((effect) => effect.kind === 'alphaModFix');
	if (alpha?.kind === 'alphaModFix' && alpha.amount < 1) paint.opacity = Math.max(0, alpha.amount);
	if (tile)
		paint.tile = {
			scaleX: tile.scaleX ?? 1,
			scaleY: tile.scaleY ?? 1,
			offsetX: (tile.offsetXEmu ?? 0) / EMU_PER_PIXEL,
			offsetY: (tile.offsetYEmu ?? 0) / EMU_PER_PIXEL,
			alignment: tile.alignment ?? 'tl',
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
}

const image = (paint: ChartPicturePaint, x: number, y: number, w: number, h: number) =>
	`<image href="${esc(paint.href)}" x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" preserveAspectRatio="none"${paint.opacity === undefined ? '' : ` opacity="${n(paint.opacity)}"`}/>`;

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
		const along = vertical
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
		return `<pattern id="${esc(id)}" patternUnits="userSpaceOnUse" x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}">${image(paint, 0, 0, w, h)}</pattern>`;
	}
	const inset = paint.inset ?? { l: 0, t: 0, r: 0, b: 0 };
	return `<pattern id="${esc(id)}" patternUnits="objectBoundingBox" patternContentUnits="objectBoundingBox" width="1" height="1">${image(paint, inset.l, inset.t, 1 - inset.l - inset.r, 1 - inset.t - inset.b)}</pattern>`;
}
