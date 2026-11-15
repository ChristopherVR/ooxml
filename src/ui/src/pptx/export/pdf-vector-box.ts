/**
 * An HTML element's own box (background colour, one linear gradient or stretched picture,
 * uniform rounded borders) as display items, or the reason it has to be drawn as a picture.
 */
import {
	cornerRadius,
	cssColor,
	linearGradientShading,
	roundedRectPath,
	splitTopLevel,
} from './pdf-vector-css';
import type { PdfExportDiagnosticCode } from './pdf-export-diagnostics';
import type { ImageItem, Matrix, PathItem, PathStroke, PdfClip } from './pdf-vector-display-list';

export interface BoxGeometry {
	x: number;
	y: number;
	width: number;
	height: number;
}

/** A background picture to decode later into `item` (already in paint order). */
export interface BoxImageRequest {
	href: string;
	fit: 'fill' | 'cover' | 'contain';
	item: ImageItem;
}

export interface BoxPaint {
	items: (PathItem | ImageItem)[];
	image?: BoxImageRequest;
	/** Set when the box cannot be vector; the caller draws it as a picture. */
	raster?: PdfExportDiagnosticCode;
}

const FILL_REASON = 'boxFillAsPicture' satisfies PdfExportDiagnosticCode;
const BORDER_REASON = 'borderAsPicture' satisfies PdfExportDiagnosticCode;
export const SHADOW_REASON = 'shadowAsPicture' satisfies PdfExportDiagnosticCode;

const px = (value: string) => Number.parseFloat(value) || 0;

/** The box's own paint; `boxShadow` is the document's shadow (editor outlines excluded). */
export function paintBox(
	style: CSSStyleDeclaration,
	box: BoxGeometry,
	matrix: Matrix,
	clips: readonly PdfClip[],
	alpha: number,
	boxShadow: string,
	/** A cell of a `border-collapse: collapse` table: its borders straddle the box edges. */
	collapsed = false,
): BoxPaint {
	if (boxShadow && boxShadow !== 'none') return { items: [], raster: SHADOW_REASON };
	const radii = [
		style.borderTopLeftRadius,
		style.borderTopRightRadius,
		style.borderBottomRightRadius,
		style.borderBottomLeftRadius,
	].map(cornerRadius);
	if (radii.some((r) => r === undefined)) return { items: [], raster: BORDER_REASON };
	const corner = radii as [number, number, number, number];
	const { x, y, width: w, height: h } = box;
	if (!(w > 0 && h > 0)) return { items: [] };
	const outline = roundedRectPath(x, y, w, h, corner);
	const base = { matrix, clips, alpha };
	const items: (PathItem | ImageItem)[] = [];
	let image: BoxImageRequest | undefined;
	const background = cssColor(style.backgroundColor);
	if (background === undefined) return { items: [], raster: FILL_REASON };
	if (background) items.push({ kind: 'path', ...base, d: outline, fill: background });
	const layers =
		style.backgroundImage && style.backgroundImage !== 'none'
			? splitTopLevel(style.backgroundImage)
			: [];
	if (layers.length > 1) return { items: [], raster: FILL_REASON };
	const layer = layers[0];
	if (layer?.startsWith('linear-gradient(')) {
		const gradient = linearGradientShading(layer, w, h);
		if (!gradient || gradient.variableOpacity) return { items: [], raster: FILL_REASON };
		items.push({
			kind: 'path',
			...base,
			d: outline,
			shading: {
				shading: gradient.shading,
				matrices: [[1, 0, 0, 1, x, y]],
				opacity: gradient.opacity,
			},
		});
	} else if (layer) {
		const href = /^url\(\s*(['"]?)([^'")]+)\1\s*\)$/.exec(layer)?.[2];
		const size = style.backgroundSize.trim();
		const stretched = size === '100% 100%' || size === 'cover' || size === 'contain';
		if (!href || !/^(data:image\/|blob:)/.test(href) || !stretched)
			return { items: [], raster: FILL_REASON };
		const item: ImageItem = {
			kind: 'image',
			matrix,
			clips: [...clips, { matrix, d: outline }],
			alpha,
			x,
			y,
			width: w,
			height: h,
			image: { width: 0, height: 0, data: new Uint8Array() },
		};
		items.push(item);
		image = {
			href,
			fit: size === 'cover' ? 'cover' : size === 'contain' ? 'contain' : 'fill',
			item,
		};
	}
	const sides = ['Top', 'Right', 'Bottom', 'Left'].map((side) => ({
		width: px(style.getPropertyValue(`border-${side.toLowerCase()}-width`)),
		style: style.getPropertyValue(`border-${side.toLowerCase()}-style`),
		color: style.getPropertyValue(`border-${side.toLowerCase()}-color`),
	}));
	const drawn = sides.filter((s) => s.width > 0 && s.style !== 'none' && s.style !== 'hidden');
	/** How far a collapsed border sits outside the box (half its width). */
	const shift = (width: number) => (collapsed ? width / 2 : 0);
	if (drawn.length) {
		const first = drawn[0]!;
		const uniform =
			drawn.length === 4 &&
			drawn.every(
				(s) => s.width === first.width && s.style === first.style && s.color === first.color,
			);
		const simple = drawn.every((s) => /^(solid|dashed|dotted)$/.test(s.style));
		if (!simple) return { items: [], raster: BORDER_REASON };
		if (uniform) {
			const color = cssColor(first.color);
			if (color === undefined) return { items: [], raster: BORDER_REASON };
			if (color) {
				const bw = first.width;
				const inset = collapsed ? 0 : bw / 2;
				const stroke: PathStroke = {
					color,
					width: bw,
					dash:
						first.style === 'dashed'
							? [bw * 3, bw * 3]
							: first.style === 'dotted'
								? [0, bw * 2]
								: [],
					cap: first.style === 'dotted' ? 1 : 0,
					join: 0,
				};
				items.push({
					kind: 'path',
					...base,
					d: roundedRectPath(
						x + inset,
						y + inset,
						w - 2 * inset,
						h - 2 * inset,
						corner.map((r) => Math.max(0, r - inset)) as [number, number, number, number],
					),
					stroke,
				});
			}
		} else {
			if (corner.some((r) => r > 0) || drawn.some((s) => s.style !== 'solid'))
				return { items: [], raster: BORDER_REASON };
			const [top, right, bottom, left] = sides;
			const rects: [typeof top, number, number, number, number][] = [
				[top, x, y - shift(top!.width), w, top!.width],
				[right, x + w - right!.width + shift(right!.width), y, right!.width, h],
				[bottom, x, y + h - bottom!.width + shift(bottom!.width), w, bottom!.width],
				[left, x - shift(left!.width), y, left!.width, h],
			];
			for (const [side, rx, ry, rw, rh] of rects) {
				if (!side || !(side.width > 0) || side.style !== 'solid') continue;
				const color = cssColor(side.color);
				if (color === undefined) return { items: [], raster: BORDER_REASON };
				if (color)
					items.push({ kind: 'path', ...base, d: roundedRectPath(rx, ry, rw, rh), fill: color });
			}
		}
	}
	return { items, ...(image ? { image } : {}) };
}
