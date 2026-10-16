/**
 * SVG markup for the chart element effects PowerPoint draws (outer shadow,
 * glow, soft edges) and for element boxes (plot area, legend, title, axis
 * label area). The markup becomes a self-contained `data:` SVG image painted
 * through a `<pattern>` def, a def kind every binding already projects, so
 * Gaussian-blurred effects reach all five bindings with no projector change.
 *
 * Geometry is in the chart's own SVG coordinates; sizes in points are scaled
 * with `CHART_PX_PER_PT`, as the chart's fonts are.
 *
 * @module render/chart-format/chart-format-effects-svg
 */
import type {
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementLine,
} from 'ooxml-core/pptx';

import { CHART_PX_PER_PT } from '../chart-font';
import type { PaintBox } from './chart-format-paint-svg';
import { gradientPaint, patternPaint, picturePaint, resetPaintIds } from './chart-format-paint-svg';

/** One shape an effect follows. */
export type EffectShape =
	| { kind: 'rect'; x: number; y: number; w: number; h: number; rx?: number }
	| { kind: 'path'; d: string; filled: boolean; strokeWidth?: number }
	| { kind: 'poly'; points: string; closed: boolean; filled: boolean; strokeWidth?: number }
	| { kind: 'circle'; cx: number; cy: number; r: number };

const r2 = (v: number) => Math.round(v * 100) / 100;
export const ptToPx = (pt: number) => pt * CHART_PX_PER_PT;
/** Opacity (0..1) of a transparency (0..100). */
export const opacityOf = (transparency: number | undefined) =>
	Math.max(0, Math.min(1, 1 - (transparency ?? 0) / 100));

/** A colour safe to put in markup: `#hex` only (model colours are resolved hex). */
export function safeColor(color: string | undefined, fallback = '#000000'): string {
	return color && /^#[0-9a-f]{3,8}$/iu.test(color) ? color : fallback;
}

const esc = (value: string) => value.replace(/[<>&"]/gu, '');

interface Paint {
	fill?: string;
	stroke?: string;
	strokeWidth?: number;
}

/** The markup of `shape` painted with `paint`. */
export function shapeMarkup(shape: EffectShape, paint: Paint): string {
	const fill = paint.fill ?? 'none';
	const stroke = paint.stroke
		? ` stroke="${paint.stroke}" stroke-width="${r2(paint.strokeWidth ?? 1)}" stroke-linejoin="round"`
		: '';
	switch (shape.kind) {
		case 'rect':
			return `<rect x="${r2(shape.x)}" y="${r2(shape.y)}" width="${r2(shape.w)}" height="${r2(shape.h)}"${shape.rx ? ` rx="${r2(shape.rx)}"` : ''} fill="${fill}"${stroke}/>`;
		case 'circle':
			return `<circle cx="${r2(shape.cx)}" cy="${r2(shape.cy)}" r="${r2(shape.r)}" fill="${fill}"${stroke}/>`;
		case 'path':
			return `<path d="${esc(shape.d)}" fill="${shape.filled ? fill : 'none'}"${stroke}/>`;
		default: {
			const tag = shape.closed ? 'polygon' : 'polyline';
			return `<${tag} points="${esc(shape.points)}" fill="${shape.filled ? fill : 'none'}"${stroke}/>`;
		}
	}
}

/** Whether `shape` is drawn as an outline only (a line series). */
const isOutline = (shape: EffectShape) =>
	(shape.kind === 'path' || shape.kind === 'poly') && !shape.filled;

let serial = 0;
/** A fresh id for a filter or mask inside one image (ids restart per image, see {@link svgDataUrl}). */
const uid = (prefix: string) => `${prefix}${++serial}`;

/** Restart the in-image ids, so the same chart always yields the same markup. */
export function resetEffectIds(): void {
	serial = 0;
	resetPaintIds();
}

/** Outer shadow: the shapes in the shadow colour, offset along `angle` and blurred. */
function shadowMarkup(
	shapes: readonly EffectShape[],
	shadow: NonNullable<PptxChartElementEffects['shadow']>,
): string {
	const id = uid('s');
	const blur = ptToPx(shadow.blur) / 2;
	const dist = ptToPx(shadow.distance);
	const rad = (shadow.angle * Math.PI) / 180;
	const color = safeColor(shadow.color);
	const body = shapes
		.map((shape) =>
			shapeMarkup(
				shape,
				isOutline(shape)
					? {
							stroke: color,
							strokeWidth:
								shape.kind !== 'rect' && shape.kind !== 'circle' ? (shape.strokeWidth ?? 1) : 1,
						}
					: { fill: color },
			),
		)
		.join('');
	return `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${r2(blur)}"/></filter><g filter="url(#${id})" opacity="${r2(opacityOf(shadow.transparency))}" transform="translate(${r2(dist * Math.cos(rad))} ${r2(dist * Math.sin(rad))})">${body}</g>`;
}

/** Glow: the shapes outlined in the glow colour (half the radius each side), blurred to fade out. */
function glowMarkup(
	shapes: readonly EffectShape[],
	glow: NonNullable<PptxChartElementEffects['glow']>,
): string {
	const id = uid('g');
	const radius = ptToPx(glow.radius);
	const color = safeColor(glow.color, '#FFFFFF');
	const body = shapes
		.map((shape) =>
			shapeMarkup(shape, {
				...(isOutline(shape) ? {} : { fill: color }),
				stroke: color,
				strokeWidth:
					radius +
					(isOutline(shape) && shape.kind !== 'rect' && shape.kind !== 'circle'
						? (shape.strokeWidth ?? 1)
						: 0),
			}),
		)
		.join('');
	return `<filter id="${id}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${r2(radius / 3)}"/></filter><g filter="url(#${id})" opacity="${r2(opacityOf(glow.transparency))}">${body}</g>`;
}

/**
 * `inner` (the shapes' own paint) faded towards their edges over `softEdge`
 * points, as PowerPoint's Soft Edges does.
 */
function softEdgeMarkup(shapes: readonly EffectShape[], softEdge: number, inner: string): string {
	const id = uid('m');
	const blurId = uid('b');
	const r = ptToPx(softEdge);
	const mask = shapes
		.map((shape) => shapeMarkup(shape, { fill: 'white', stroke: 'black', strokeWidth: r }))
		.join('');
	return `<filter id="${blurId}" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="${r2(r / 4)}"/></filter><mask id="${id}" maskUnits="userSpaceOnUse"><g filter="url(#${blurId})">${mask}</g></mask><g mask="url(#${id})">${inner}</g>`;
}

/** Effects behind a shape group (shadow, glow), as markup. */
export function effectsBehind(
	shapes: readonly EffectShape[],
	effects: PptxChartElementEffects | undefined,
): string {
	if (!effects || shapes.length === 0) {
		return '';
	}
	return (
		(effects.shadow ? shadowMarkup(shapes, effects.shadow) : '') +
		(effects.glow && effects.glow.radius > 0 ? glowMarkup(shapes, effects.glow) : '')
	);
}

/** `fill` paint of a shape group, faded at its edges when `softEdge` is set. */
export function paintedShapes(
	shapes: readonly EffectShape[],
	paint: (shape: EffectShape, index: number) => string,
	softEdge: number | null | undefined,
): string {
	const inner = shapes.map(paint).join('');
	return softEdge && softEdge > 0 ? softEdgeMarkup(shapes, softEdge, inner) : inner;
}

/** The paint server and opacity of a fill in `box`, or `undefined` when it paints nothing. */
export function fillPaintOf(
	fill: PptxChartElementFill | undefined,
	box: PaintBox,
	fallbackFill?: string,
): { defs: string; paint: string; opacity: number } | undefined {
	switch (fill?.kind) {
		case 'solid':
			return {
				defs: '',
				paint: safeColor(fill.color, fallbackFill ?? '#FFFFFF'),
				opacity: opacityOf(fill.transparency),
			};
		case 'gradient': {
			if (!fill.gradient || fill.gradient.stops.length < 2) return undefined;
			const server = gradientPaint(fill.gradient, box);
			return { ...server, opacity: 1 };
		}
		case 'pattern': {
			const server = fill.pattern ? patternPaint(fill.pattern) : undefined;
			return server ? { ...server, opacity: 1 } : undefined;
		}
		case 'picture': {
			const src = fill.picture?.src;
			if (!src) return undefined;
			const server = picturePaint(src, box, Boolean(fill.picture?.tile));
			return { ...server, opacity: opacityOf(fill.picture?.transparency) };
		}
		case 'none':
			return undefined;
		default:
			return fallbackFill ? { defs: '', paint: fallbackFill, opacity: 1 } : undefined;
	}
}

/** The stroke paint of a line over `box` (solid or gradient), or `undefined`. */
export function linePaintOf(
	line: PptxChartElementLine | undefined,
	box: PaintBox,
): { defs: string; paint: string; opacity: number } | undefined {
	if (line?.kind === 'solid') {
		return { defs: '', paint: safeColor(line.color), opacity: opacityOf(line.transparency) };
	}
	if (line?.kind === 'gradient' && line.gradient && line.gradient.stops.length >= 2) {
		return { ...gradientPaint(line.gradient, box), opacity: opacityOf(line.transparency) };
	}
	return undefined;
}

/** Stroke widths and offsets (fractions of the width, from the centre line) of a compound line. */
export function compoundStrokes(
	compound: PptxChartElementLine['compound'],
): Array<[number, number]> {
	switch (compound) {
		case 'dbl':
			return [
				[1 / 3, -1 / 3],
				[1 / 3, 1 / 3],
			];
		case 'thickThin':
			return [
				[0.6, -0.2],
				[0.2, 0.4],
			];
		case 'thinThick':
			return [
				[0.2, -0.4],
				[0.6, 0.2],
			];
		case 'tri':
			return [
				[0.2, -0.4],
				[0.2, 0],
				[0.2, 0.4],
			];
		default:
			return [[1, 0]];
	}
}

/**
 * A box (plot area, legend, title, axis label area) with its fill, border
 * and effects. `fallbackFill` is the box's flat-field fill when the format
 * leaves the fill to it.
 */
export function boxMarkup(
	box: { x: number; y: number; w: number; h: number; rx?: number },
	fill: PptxChartElementFill | undefined,
	line: PptxChartElementLine | undefined,
	effects: PptxChartElementEffects | undefined,
	fallbackFill?: string,
): string {
	const shape: EffectShape = { kind: 'rect', ...box };
	const paint = fillPaintOf(fill, box, fallbackFill);
	const body = paint
		? paint.defs +
			paintedShapes(
				[shape],
				(s) =>
					shapeMarkup(s, { fill: paint.paint }).replace(
						'/>',
						` fill-opacity="${r2(paint.opacity)}"/>`,
					),
				effects?.softEdge,
			)
		: '';
	const stroke = linePaintOf(line, box);
	const width = ptToPx(line?.width ?? 0.75);
	const border = stroke
		? stroke.defs +
			compoundStrokes(line?.compound)
				.map(([share, offset]) =>
					shapeMarkup(
						{
							kind: 'rect',
							x: box.x - offset * width,
							y: box.y - offset * width,
							w: box.w + 2 * offset * width,
							h: box.h + 2 * offset * width,
							...(box.rx ? { rx: box.rx } : {}),
						},
						{ stroke: stroke.paint, strokeWidth: width * share },
					).replace('/>', ` stroke-opacity="${r2(stroke.opacity)}"/>`),
				)
				.join('')
		: '';
	const casting = paint || border ? effectsBehind([shape], effects) : '';
	return casting + body + border;
}

/** A complete SVG document of `width` x `height` (chart coordinates) as a `data:` URL. */
export function svgDataUrl(width: number, height: number, body: string): string {
	const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${r2(width)}" height="${r2(height)}" viewBox="0 0 ${r2(width)} ${r2(height)}">${body}</svg>`;
	return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
}
