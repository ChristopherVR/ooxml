/**
 * Draws the chart element formats (`PptxChartData.elementFormats`, edited in
 * the Format panes or read from each element's `c:spPr`) that the flat chart
 * fields cannot express: fill and line transparency, outer shadow, glow,
 * soft edges, and the plot area / legend / title / axis label boxes.
 *
 * Everything lands in parts of the view model every binding already
 * projects: primitive opacity, a data-label box colour, and `<pattern>` defs
 * whose image is a self-contained SVG (`chart-format-effects-svg.ts`).
 * The boxes are composed into the chart-area background, so they paint
 * under the gridlines, labels and marks; series effects are a pattern-filled
 * rectangle placed just before that series' marks.
 *
 * @module render/chart-format/chart-format-render
 */
import type {
	PptxChartData,
	PptxChartDataLabelOptions,
	PptxChartElementFormat,
	PptxChartShapeProps,
} from 'ooxml-core/pptx';

import { computeChartLegendLayout } from '../chart-legend-layout';
import type { ChartSvgPatternDef } from '../chart-svg-def-types';
import type { ChartViewModel, SvgPrimitive, SvgRect, SvgText } from '../chart-view-model-types';
import type { EffectShape } from './chart-format-effects-svg';
import {
	boxMarkup,
	effectsBehind,
	opacityOf,
	paintedShapes,
	resetEffectIds,
	safeColor,
	shapeMarkup,
	svgDataUrl,
} from './chart-format-effects-svg';
import { chartAxisIndex, chartPlotBounds, chartTextBox } from './chart-format-geometry';

const HEX = /^#[0-9a-f]{6}$/iu;

/** `#RRGGBB` at `opacity`, as an `rgba()` the SVG `fill` attribute accepts. */
function withOpacity(color: string, opacity: number): string {
	if (opacity >= 1 || !HEX.test(color)) {
		return color;
	}
	const n = Number.parseInt(color.slice(1), 16);
	return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${Math.round(opacity * 1000) / 1000})`;
}

/** The label box a data-label format paints, over the parsed one. */
function labelShapeFrom(
	format: PptxChartElementFormat | undefined,
	current: PptxChartShapeProps | undefined,
): PptxChartShapeProps | undefined {
	const fill = format?.fill;
	const line = format?.line;
	if (fill?.kind !== 'solid' && line?.kind !== 'solid') {
		return current;
	}
	const next: PptxChartShapeProps = { ...current };
	if (fill?.kind === 'solid' && fill.color) {
		next.fillColor = withOpacity(safeColor(fill.color), opacityOf(fill.transparency));
	} else if (fill?.kind === 'none') {
		next.fillColor = 'none';
	}
	if (line?.kind === 'solid' && line.color) {
		next.strokeColor = withOpacity(safeColor(line.color), opacityOf(line.transparency));
		if (line.width !== undefined) next.strokeWidth = line.width;
	} else if (line?.kind === 'none') {
		next.strokeColor = 'none';
	}
	return next;
}

/**
 * Render-only copy of `data` whose data-label boxes follow the data-label
 * formats (colour, transparency, border), for the label decorations.
 */
export function withElementFormatLabelBoxes(data: PptxChartData): PptxChartData {
	const formats = data.elementFormats;
	if (!formats) {
		return data;
	}
	const chartWide = formats.dataLabels;
	const withShape = (
		opts: PptxChartDataLabelOptions | undefined,
		format: PptxChartElementFormat | undefined,
	) => {
		const shape = labelShapeFrom(format, opts?.labelShape);
		return shape === opts?.labelShape ? opts : { ...opts, labelShape: shape };
	};
	const style = data.style?.dataLabels
		? { ...data.style, dataLabels: withShape(data.style.dataLabels, chartWide) }
		: data.style;
	let changed = style !== data.style;
	const series = data.series.map((s, i) => {
		const format = formats[`dataLabels:${i}`] ?? chartWide;
		if (!s.dataLabelOptions || !format) return s;
		const opts = withShape(s.dataLabelOptions, format);
		if (opts === s.dataLabelOptions) return s;
		changed = true;
		return { ...s, dataLabelOptions: opts };
	});
	return changed ? { ...data, ...(style ? { style } : {}), series } : data;
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

interface Images {
	defs: ChartSvgPatternDef[];
	idFor: (suffix: string) => string;
	width: number;
	height: number;
}

/** A full-chart rectangle painted with an SVG image of `body`. */
function imageRect(images: Images, suffix: string, body: string): SvgRect {
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

/** Series transparency and effects, applied to the series' own marks. */
function applySeriesFormats(
	vm: ChartViewModel,
	data: PptxChartData,
	images: Images,
): SvgPrimitive[] {
	let primitives = vm.primitives;
	data.series.forEach((_, si) => {
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
	return primitives;
}

const hasBox = (format: PptxChartElementFormat | undefined) =>
	Boolean(
		format &&
		!format.reset &&
		(format.fill?.kind === 'solid' ||
			format.line?.kind === 'solid' ||
			format.effects?.shadow ||
			(format.effects?.glow && format.effects.glow.radius > 0)),
	);

function union(
	boxes: Array<{ left: number; top: number; right: number; bottom: number }>,
	pad: number,
) {
	if (boxes.length === 0) return undefined;
	const left = Math.min(...boxes.map((b) => b.left)) - pad;
	const top = Math.min(...boxes.map((b) => b.top)) - pad;
	const right = Math.max(...boxes.map((b) => b.right)) + pad;
	const bottom = Math.max(...boxes.map((b) => b.bottom)) + pad;
	return { x: left, y: top, w: right - left, h: bottom - top };
}

const textBoxes = (texts: readonly SvgText[] | undefined) =>
	(texts ?? []).filter((t) => !t.transform).map((t) => chartTextBox(t, 2));

/** The chart-area background with the element boxes, or `undefined` when plain. */
function backgroundMarkup(vm: ChartViewModel, data: PptxChartData): string | undefined {
	const f = data.elementFormats ?? {};
	const layers: string[] = [];
	const plot = f.plotArea;
	const plotBox =
		hasBox(plot) || (plot?.effects?.softEdge ?? 0) > 0 ? chartPlotBounds(vm) : undefined;
	if (plot && plotBox) {
		layers.push(
			boxMarkup(
				{
					x: plotBox.left,
					y: plotBox.top,
					w: plotBox.right - plotBox.left,
					h: plotBox.bottom - plotBox.top,
				},
				plot.fill,
				plot.line,
				plot.effects,
			),
		);
	}
	const axisSources: Array<[number, readonly SvgText[] | undefined]> = [
		[chartAxisIndex(data, ['catAx', 'dateAx']), vm.categoryLabels],
		[chartAxisIndex(data, ['valAx']), vm.axisLabels],
		[chartAxisIndex(data, ['valAx'], 1), vm.secondaryAxisLabels],
	];
	for (const [index, texts] of axisSources) {
		const format = index >= 0 ? f[`axis:${index}`] : undefined;
		// An axis' line is the axis line itself (drawn with the axis); its fill paints the label area.
		if (format?.fill?.kind === 'solid' && !format.reset) {
			const box = union(textBoxes(texts), 2);
			if (box) layers.push(boxMarkup(box, format.fill, undefined, format.effects));
		}
	}
	const legend = f.legend;
	if (hasBox(legend) && data.style?.hasLegend !== false && vm.legend.length > 0) {
		const items = computeChartLegendLayout(vm).map((item) => ({
			left: item.x,
			top: item.y - 8,
			right: item.x + 14 + item.label.length * item.fontSize * 0.58,
			bottom: item.y + 5,
		}));
		const box = union(items, 4);
		if (box) layers.push(boxMarkup(box, legend!.fill, legend!.line, legend!.effects));
	}
	const title = f.title;
	if (hasBox(title) && vm.title) {
		const size = vm.titleStyle?.fontSize ?? 12;
		const box = union(
			[
				chartTextBox(
					{
						kind: 'text',
						x: vm.titleX,
						y: vm.titleY,
						text: vm.title,
						fontSize: size,
						fill: '',
						textAnchor: 'middle',
					},
					3,
				),
			],
			3,
		);
		if (box) layers.push(boxMarkup(box, title!.fill, title!.line, title!.effects));
	}
	const area = f.chartArea;
	const areaSpecial =
		area &&
		!area.reset &&
		((area.fill?.kind === 'solid' && (area.fill.transparency ?? 0) > 0) ||
			(area.effects?.softEdge ?? 0) > 0);
	if (layers.length === 0 && !areaSpecial) {
		return undefined;
	}
	const base = vm.areaFill && HEX.test(vm.areaFill) ? vm.areaFill : undefined;
	const areaLayer = base
		? boxMarkup(
				{
					x: 0,
					y: 0,
					w: vm.svgWidth,
					h: vm.svgHeight,
					...(vm.areaRadius ? { rx: vm.areaRadius } : {}),
				},
				area?.fill?.kind === 'solid' ? { ...area.fill, color: base } : undefined,
				undefined,
				area?.effects?.softEdge ? { softEdge: area.effects.softEdge } : undefined,
				base,
			)
		: '';
	return areaLayer + layers.join('');
}

/**
 * Apply the element formats of `data` to a finished view model. A chart
 * without formats beyond what the flat fields draw comes back unchanged.
 */
export function applyChartElementFormats(
	vm: ChartViewModel,
	data: PptxChartData,
	elementId: string,
): ChartViewModel {
	if (!data.elementFormats) {
		return vm;
	}
	const images: Images = {
		defs: [],
		idFor: (suffix) => `${elementId}-fmt-${suffix}`.replace(/[^A-Za-z0-9_-]/gu, '_'),
		width: vm.svgWidth,
		height: vm.svgHeight,
	};
	const primitives = applySeriesFormats(vm, data, images);
	let next: ChartViewModel = primitives === vm.primitives ? vm : { ...vm, primitives };
	// A gradient chart area keeps its own def; the boxes then paint as the first primitive.
	resetEffectIds();
	const background = backgroundMarkup(next, data);
	if (background) {
		const gradientArea = Boolean(next.areaFill && !HEX.test(next.areaFill));
		const rect = imageRect(images, 'bg', background);
		next = gradientArea
			? { ...next, primitives: [rect, ...next.primitives] }
			: { ...next, areaFill: rect.fill };
	}
	return images.defs.length > 0 ? { ...next, defs: [...(next.defs ?? []), ...images.defs] } : next;
}
