/**
 * An approximate drawing of a chart element's 3-D Format, as markup for the
 * effect images: PowerPoint renders bevels with real lighting, which a flat
 * SVG cannot. Seen from the front (every chart element's default camera) a
 * top bevel shows as light and shadow bands along the shape's edges and a
 * contour as an outline; the material and lighting rig only scale how strong
 * the bands are. Depth and the bottom bevel face away from the viewer and
 * are not drawn.
 *
 * @module render/chart-format/chart-format-3d-svg
 */
import type { PptxChartElementThreeD } from 'ooxml-core/pptx';

import { ptToPx, safeColor } from './chart-format-effects-svg';

export interface BevelRect {
	x: number;
	y: number;
	w: number;
	h: number;
}

const r2 = (v: number) => Math.round(v * 100) / 100;

/** Highlight and shade strengths (0..1) a material gives a bevel. */
const MATERIAL: Record<string, [number, number]> = {
	warmMatte: [0.45, 0.3],
	matte: [0.35, 0.25],
	plastic: [0.55, 0.3],
	metal: [0.7, 0.45],
	dkEdge: [0.35, 0.55],
	softEdge: [0.3, 0.2],
	flat: [0.12, 0.08],
	softmetal: [0.55, 0.35],
	clear: [0.3, 0.15],
	powder: [0.3, 0.2],
	translucentPowder: [0.25, 0.15],
	legacyMatte: [0.35, 0.25],
	legacyPlastic: [0.55, 0.3],
	legacyMetal: [0.7, 0.45],
	legacyWireframe: [0.1, 0.1],
};
const LIGHTING: Record<string, number> = { harsh: 1.3, soft: 0.7, flat: 0.35, flood: 1.1 };
/** Bevel shapes with a hard ridge rather than a rounded one. */
const HARD = new Set(['angle', 'slope', 'hardEdge', 'artDeco', 'cross', 'divot', 'riblet']);

let serial = 0;
/** Restart the ids inside one image. */
export function resetBevelIds(): void {
	serial = 0;
}

function band(
	id: string,
	rect: BevelRect,
	side: 'top' | 'left' | 'bottom' | 'right',
	size: number,
	color: string,
	opacity: number,
	hard: boolean,
): string {
	const vertical = side === 'top' || side === 'bottom';
	const [x1, y1, x2, y2] =
		side === 'top'
			? [0, 0, 0, 1]
			: side === 'bottom'
				? [0, 1, 0, 0]
				: side === 'left'
					? [0, 0, 1, 0]
					: [1, 0, 0, 0];
	const stops = hard
		? `<stop offset="0" stop-color="${color}" stop-opacity="${r2(opacity)}"/><stop offset="0.85" stop-color="${color}" stop-opacity="${r2(opacity * 0.8)}"/><stop offset="1" stop-color="${color}" stop-opacity="0"/>`
		: `<stop offset="0" stop-color="${color}" stop-opacity="${r2(opacity)}"/><stop offset="1" stop-color="${color}" stop-opacity="0"/>`;
	const bx = side === 'right' ? rect.x + rect.w - size : rect.x;
	const by = side === 'bottom' ? rect.y + rect.h - size : rect.y;
	const bw = vertical ? rect.w : size;
	const bh = vertical ? size : rect.h;
	return `<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}">${stops}</linearGradient><rect x="${r2(bx)}" y="${r2(by)}" width="${r2(bw)}" height="${r2(bh)}" fill="url(#${id})"/>`;
}

/** The bevel and contour of `threeD` over `rects`, or `''` when nothing shows from the front. */
export function bevelMarkup(
	rects: readonly BevelRect[],
	threeD: PptxChartElementThreeD | undefined,
): string {
	if (!threeD) {
		return '';
	}
	const bevel = threeD.bevelTop ?? undefined;
	const contour = threeD.contour && threeD.contour.size > 0 ? threeD.contour : undefined;
	if (!bevel && !contour) {
		return '';
	}
	const [light, dark] = MATERIAL[threeD.material ?? 'warmMatte'] ?? MATERIAL.warmMatte!;
	const scale = LIGHTING[threeD.lighting ?? 'threePt'] ?? 1;
	const hard = bevel ? HARD.has(bevel.type) : false;
	return rects
		.filter((rect) => rect.w > 0 && rect.h > 0)
		.map((rect) => {
			let out = '';
			if (bevel && bevel.width > 0) {
				const size = Math.min(ptToPx(bevel.width), rect.w / 2, rect.h / 2);
				const strength = Math.min(
					1,
					(bevel.height > 0 ? 0.6 + Math.min(bevel.height, 12) / 30 : 0.6) * scale,
				);
				const id = () => `bv${++serial}`;
				out +=
					band(id(), rect, 'top', size, '#FFFFFF', light * strength, hard) +
					band(id(), rect, 'left', size, '#FFFFFF', light * strength * 0.6, hard) +
					band(id(), rect, 'bottom', size, '#000000', dark * strength, hard) +
					band(id(), rect, 'right', size, '#000000', dark * strength * 0.6, hard);
			}
			if (contour) {
				const width = ptToPx(contour.size);
				out += `<rect x="${r2(rect.x - width / 2)}" y="${r2(rect.y - width / 2)}" width="${r2(rect.w + width)}" height="${r2(rect.h + width)}" fill="none" stroke="${safeColor(contour.color, '#FFFFFF')}" stroke-width="${r2(width)}"/>`;
			}
			return out;
		})
		.join('');
}
