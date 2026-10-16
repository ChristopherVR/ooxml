/**
 * Series formats the flat series fields cannot draw: fill and line
 * transparency, outer shadow, glow and soft edges, pattern fills (a
 * `<pattern>` def of the preset tile, so the marks and the legend swatch
 * paint with it) and the 3-D Format bevel and contour of rectangular marks.
 *
 * @module render/chart-format/chart-format-render-series
 */
import type { PptxChartData } from 'ooxml-core/pptx';

import type { ChartSvgPatternDef } from '../chart-svg-def-types';
import type { ChartViewModel, SvgPrimitive, SvgRect } from '../chart-view-model-types';
import { getPatternTile } from '../fill-style';
import { bevelMarkup, resetBevelIds } from './chart-format-3d-svg';
import type { EffectShape } from './chart-format-effects-svg';
import {
	effectsBehind,
	opacityOf,
	paintedShapes,
	resetEffectIds,
	shapeMarkup,
	svgDataUrl,
} from './chart-format-effects-svg';

const HEX = /^#[0-9a-f]{6}$/iu;

/** The pattern defs and full-chart effect images one chart's formats add. */
export interface Images {
	defs: ChartSvgPatternDef[];
	idFor: (suffix: string) => string;
	width: number;
	height: number;
}

/** A full-chart rectangle painted with an SVG image of `body`. */
export function imageRect(images: Images, suffix: string, body: string): SvgRect {
	const id = images.idFor(suffix);
	images.defs.push({
		kind: 'pattern',
		id,
		href: svgDataUrl(images.width, images.height, body),
		patternUnits: 'userSpaceOnUse',
		x: 0,
		y: 0,
		width: images.width,
		height: images.height,
	});
	return { kind: 'rect', x: 0, y: 0, w: images.width, h: images.height, fill: `url(#${id})` };
}

function shapeOf(prim: SvgPrimitive): EffectShape | undefined {
	switch (prim.kind) {
		case 'rect':
			return {
				kind: 'rect',
				x: prim.x,
				y: prim.y,
				w: prim.w,
				h: prim.h,
				...(prim.rx ? { rx: prim.rx } : {}),
			};
		case 'circle':
			return { kind: 'circle', cx: prim.cx, cy: prim.cy, r: prim.r };
		case 'path':
			return {
				kind: 'path',
				d: prim.d,
				filled: prim.fill !== 'none',
				strokeWidth: prim.strokeWidth ?? 1,
			};
		case 'polyline':
			return {
				kind: 'poly',
				points: prim.points,
				closed: false,
				filled: false,
				strokeWidth: prim.strokeWidth,
			};
		case 'polygon':
			return {
				kind: 'poly',
				points: prim.points,
				closed: true,
				filled: prim.fill !== 'none',
				strokeWidth: prim.strokeWidth,
			};
		default:
			return undefined;
	}
}

const isFilled = (prim: SvgPrimitive) =>
	(prim.kind === 'rect' ||
		prim.kind === 'circle' ||
		prim.kind === 'path' ||
		prim.kind === 'polygon') &&
	prim.fill !== 'none';
const isStroked = (prim: SvgPrimitive) =>
	prim.kind === 'polyline' || (prim.kind === 'path' && prim.fill === 'none');

/** A series pattern fill as a `<pattern>` def of its preset tile; returns the paint. */
function patternDef(
	images: Images,
	si: number,
	pattern: { preset: string; foreground: string; background: string },
): string | undefined {
	const tile = getPatternTile(pattern.preset, pattern.foreground, pattern.background);
	if (!tile) return undefined;
	const id = images.idFor(`pt${si}`);
	images.defs.push({
		kind: 'pattern',
		id,
		href: svgDataUrl(tile.w, tile.h, tile.inner),
		patternUnits: 'userSpaceOnUse',
		x: 0,
		y: 0,
		width: tile.w,
		height: tile.h,
	});
	return `url(#${id})`;
}

/** The bevel / contour image of a series' rectangular marks, inserted after its last mark. */
function withBevel(
	primitives: SvgPrimitive[],
	mine: (prim: SvgPrimitive) => boolean,
	threeD: NonNullable<PptxChartData['elementFormats']>[string]['threeD'],
	images: Images,
	si: number,
): SvgPrimitive[] {
	const rects = primitives.filter((prim): prim is SvgRect => mine(prim) && prim.kind === 'rect');
	resetBevelIds();
	const body = bevelMarkup(rects, threeD);
	if (!body) return primitives;
	let last = -1;
	primitives.forEach((prim, i) => {
		if (mine(prim)) last = i;
	});
	return [
		...primitives.slice(0, last + 1),
		imageRect(images, `b${si}`, body),
		...primitives.slice(last + 1),
	];
}

/**
 * Series transparency, pattern fills, effects and bevels, applied to the
 * series' own marks. Returns the new primitives and the legend swatch
 * paints that changed (by legend entry index).
 */
export function applySeriesFormats(
	vm: ChartViewModel,
	data: PptxChartData,
	images: Images,
): { primitives: SvgPrimitive[]; legendFills: Map<number, string> } {
	let primitives = vm.primitives;
	const legendFills = new Map<number, string>();
	data.series.forEach((series, si) => {
		const format = data.elementFormats?.[`series:${si}`];
		if (!format || format.reset) return;
		const fillOpacity = format.fill?.kind === 'solid' ? opacityOf(format.fill.transparency) : 1;
		const lineOpacity = format.line?.kind === 'solid' ? opacityOf(format.line.transparency) : 1;
		const mine = (prim: SvgPrimitive) => 'part' in prim && prim.part?.seriesIndex === si;
		if (fillOpacity < 1 || lineOpacity < 1) {
			primitives = primitives.map((prim) => {
				if (!mine(prim)) return prim;
				const factor = isFilled(prim) ? fillOpacity : isStroked(prim) ? lineOpacity : 1;
				return factor < 1
					? ({
							...prim,
							opacity: ((prim as { opacity?: number }).opacity ?? 1) * factor,
						} as SvgPrimitive)
					: prim;
			});
		}
		if (format.fill?.kind === 'pattern' && format.fill.pattern) {
			const paint = patternDef(images, si, format.fill.pattern);
			if (paint) {
				primitives = primitives.map((prim) =>
					mine(prim) && isFilled(prim) ? ({ ...prim, fill: paint } as SvgPrimitive) : prim,
				);
				// The legend lists series in order; its swatch follows the series' fill.
				if (vm.legend[si]?.label === series.name) legendFills.set(si, paint);
			}
		}
		if (format.threeD) {
			primitives = withBevel(primitives, mine, format.threeD, images, si);
		}
		const effects = format.effects;
		const soft = effects?.softEdge && effects.softEdge > 0 ? effects.softEdge : 0;
		if (!effects || (!effects.shadow && !(effects.glow && effects.glow.radius > 0) && !soft))
			return;
		const marks = primitives.filter(mine);
		const shapes = marks.map(shapeOf).filter((s): s is EffectShape => Boolean(s));
		if (shapes.length === 0) return;
		resetEffectIds();
		// Soft edges repaint the solid marks faded at their edges; the originals stay for hit-testing.
		const softMarks = soft
			? marks.filter((m) => isFilled(m) && HEX.test((m as { fill: string }).fill))
			: [];
		const softBody = softMarks.length
			? paintedShapes(
					softMarks.map((m) => shapeOf(m)!),
					(shape, index) =>
						shapeMarkup(shape, { fill: (softMarks[index] as { fill: string }).fill }),
					soft,
				)
			: '';
		const body = effectsBehind(shapes, effects) + softBody;
		if (!body) return;
		const first = primitives.findIndex(mine);
		const image = imageRect(images, `s${si}`, body);
		const hidden = new Set(softMarks);
		primitives = [
			...primitives.slice(0, first),
			image,
			...primitives
				.slice(first)
				.map((prim) => (hidden.has(prim) ? ({ ...prim, opacity: 0 } as SvgPrimitive) : prim)),
		];
	});
	return { primitives, legendFills };
}
