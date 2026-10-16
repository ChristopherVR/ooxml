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
import type { PptxChartData, PptxChartElementFill, PptxChartElementFormat } from 'ooxml-core/pptx';

import { computeChartLegendLayout } from '../chart-legend-layout';
import type { ChartViewModel, SvgText } from '../chart-view-model-types';
import { bevelMarkup, resetBevelIds } from './chart-format-3d-svg';
import { boxMarkup, resetEffectIds } from './chart-format-effects-svg';
import { chartAxisIndex, chartPlotBounds, chartTextBox } from './chart-format-geometry';
import { decorateAxisLines } from './chart-format-render-lines';
import type { Images } from './chart-format-render-series';
import { applySeriesFormats, imageRect } from './chart-format-render-series';
import {
	dataLabelEffectsMarkup,
	paints,
	withElementFormatLabelBoxes,
	withLabelTextEffects,
} from './chart-format-render-labels';
import { applyTextFormats } from './chart-format-render-text';

export { withElementFormatLabelBoxes };

const HEX = /^#[0-9a-f]{6}$/iu;

const hasBox = (format: PptxChartElementFormat | undefined, ownGradient = true) =>
	Boolean(
		format &&
		!format.reset &&
		(paints(format.fill, ownGradient) ||
			format.line?.kind === 'solid' ||
			format.line?.kind === 'gradient' ||
			format.effects?.shadow ||
			(format.effects?.glow && format.effects.glow.radius > 0) ||
			format.threeD?.bevelTop ||
			format.threeD?.contour),
	);

/** A box's fill as drawn here: a gradient the model paints itself is left to it. */
const boxFill = (fill: PptxChartElementFill | undefined, ownGradient: boolean) =>
	fill?.kind === 'gradient' && !ownGradient ? undefined : fill;

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

/** A box with its fill, border, effects and 3-D bevel. */
function boxLayer(
	box: { x: number; y: number; w: number; h: number; rx?: number },
	format: PptxChartElementFormat,
	ownGradient = true,
): string {
	resetBevelIds();
	return (
		boxMarkup(box, boxFill(format.fill, ownGradient), format.line, format.effects) +
		bevelMarkup([box], format.threeD)
	);
}

/** The chart-area background with the element boxes, or `undefined` when plain. */
function backgroundMarkup(vm: ChartViewModel, data: PptxChartData): string | undefined {
	const f = data.elementFormats ?? {};
	const layers: string[] = [];
	const plot = f.plotArea;
	const plotBox =
		hasBox(plot, false) || (plot?.effects?.softEdge ?? 0) > 0 ? chartPlotBounds(vm) : undefined;
	if (plot && plotBox) {
		layers.push(
			boxLayer(
				{
					x: plotBox.left,
					y: plotBox.top,
					w: plotBox.right - plotBox.left,
					h: plotBox.bottom - plotBox.top,
				},
				plot,
				false,
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
		if (format && !format.reset && paints(format.fill, true)) {
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
		if (box) layers.push(boxLayer(box, legend!));
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
		if (box) layers.push(boxLayer(box, title!));
	}
	const area = f.chartArea && !f.chartArea.reset ? f.chartArea : undefined;
	const areaPaints = area?.fill?.kind === 'pattern' || area?.fill?.kind === 'picture';
	const areaSpecial =
		area &&
		((area.fill?.kind === 'solid' && (area.fill.transparency ?? 0) > 0) ||
			areaPaints ||
			(area.effects?.softEdge ?? 0) > 0 ||
			area.threeD?.bevelTop ||
			area.threeD?.contour);
	if (layers.length === 0 && !areaSpecial) {
		return undefined;
	}
	const frame = {
		x: 0,
		y: 0,
		w: vm.svgWidth,
		h: vm.svgHeight,
		...(vm.areaRadius ? { rx: vm.areaRadius } : {}),
	};
	const base = vm.areaFill && HEX.test(vm.areaFill) ? vm.areaFill : undefined;
	resetBevelIds();
	const areaLayer =
		base || areaPaints
			? boxMarkup(
					frame,
					areaPaints
						? area!.fill
						: area?.fill?.kind === 'solid'
							? { ...area.fill, color: base! }
							: undefined,
					undefined,
					area?.effects?.softEdge ? { softEdge: area.effects.softEdge } : undefined,
					base,
				) + bevelMarkup([frame], area?.threeD)
			: '';
	return areaLayer + layers.join('');
}

/**
 * Apply the element formats of `data` to a finished view model, and strip
 * the build-time axis-line tags. A chart without formats beyond what the
 * flat fields draw comes back unchanged.
 */
export function applyChartElementFormats(
	vm: ChartViewModel,
	data: PptxChartData,
	elementId: string,
): ChartViewModel {
	const lines = decorateAxisLines(vm, data);
	if (!data.elementFormats) {
		return lines.vm;
	}
	const images: Images = {
		defs: [],
		idFor: (suffix) => `${elementId}-fmt-${suffix}`.replace(/[^A-Za-z0-9_-]/gu, '_'),
		width: vm.svgWidth,
		height: vm.svgHeight,
	};
	const series = applySeriesFormats(lines.vm, data, images);
	let next: ChartViewModel =
		series.primitives === lines.vm.primitives
			? lines.vm
			: { ...lines.vm, primitives: series.primitives };
	if (series.legendFills.size > 0) {
		next = {
			...next,
			legend: next.legend.map((entry, i) => {
				const fill = series.legendFills.get(i);
				return fill ? { ...entry, color: fill } : entry;
			}),
		};
	}
	// Decorated axis lines paint above the gridlines and below the marks.
	resetEffectIds();
	if (lines.markup) {
		next = { ...next, primitives: [imageRect(images, 'ax', lines.markup), ...next.primitives] };
	}
	// Text effects and label boxes' shadows paint above the marks, under the text nodes.
	resetEffectIds();
	const text = applyTextFormats(next, withLabelTextEffects(data));
	next = text.vm;
	resetEffectIds();
	const overlay = dataLabelEffectsMarkup(next, data) + text.markup;
	if (overlay) {
		next = { ...next, primitives: [...next.primitives, imageRect(images, 'tx', overlay)] };
	}
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
