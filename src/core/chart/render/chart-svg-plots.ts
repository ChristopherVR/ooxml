import { boxWhiskerSvg } from './chart-svg-box';
import { chartDataTableSvg } from './chart-svg-data-table';
import type { ChartFrameSize } from '../manual-layout';
// Which plot painter draws a chart view model: bars, lines and areas (alone or layered in a
// combination chart), scatter, waterfall and histogram.
import { categoryAnalysis } from './chart-svg-analysis';
import { bars } from './chart-svg-bars';
import { cartesianAxes, linesAndAreas, plotted, type Frame } from './chart-svg-cartesian';
import { scatter } from './chart-svg-xy';
import { stockBars } from './chart-svg-stock';
import { waterfallBars } from './chart-svg-waterfall';
import type { Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

interface Layer {
	model: ChartViewModel;
	frame: Frame;
}

/** Areas sit behind bars, bars behind lines, as Office stacks the groups of a combination. */
const RANK: Record<string, number> = {
	area: 0,
	bar: 1,
	column: 1,
	histogram: 1,
	line: 2,
	stock: 2,
};

function paint(layer: Layer, out: string[]): void {
	const { model, frame } = layer;
	if (model.type === 'stock') stockBars(model, frame, out);
	else if (model.type === 'bar' || model.type === 'column' || model.type === 'histogram')
		bars(model, frame, out, plotted(model));
	else linesAndAreas(model, frame, out);
	categoryAnalysis(model, frame, out);
}

/** A split group still clusters columns on the same scale together, retaining series indices. */
function clusterLayers(layers: Layer[]): Layer[] {
	const out: Layer[] = [];
	for (const layer of layers) {
		const match =
			layer.model.type === 'column' && layer.model.grouping === 'clustered'
				? out.find(
						(other) =>
							other.model.type === 'column' &&
							other.model.grouping === 'clustered' &&
							other.frame.value === layer.frame.value &&
							other.model.barGapWidth === layer.model.barGapWidth &&
							other.model.barOverlap === layer.model.barOverlap,
					)
				: undefined;
		if (match) match.model.series = [...match.model.series, ...layer.model.series];
		else out.push({ ...layer, model: { ...layer.model, series: [...layer.model.series] } });
	}
	return out;
}

/** Paints the plot types that share cartesian axes into `area`. */
export function cartesianSvg(model: ChartViewModel, area: Rect, size?: ChartFrameSize): string {
	const out: string[] = [];
	if (model.type === 'scatter' || model.type === 'bubble') {
		scatter(model, area, out, size);
		return out.join('');
	}
	const frame = cartesianAxes(model, area, out, size);
	if (!frame) return '';
	out.push(chartDataTableSvg(model, frame.plot));
	if (model.type === 'boxWhisker') {
		boxWhiskerSvg(model, frame, out);
		return out.join('');
	}
	if (model.type === 'waterfall') {
		waterfallBars(model, frame, out);
		return out.join('');
	}
	if (model.type === 'stock') {
		stockBars(model, frame, out);
		return out.join('');
	}
	const layers: Layer[] = [{ model, frame }];
	for (const group of model.extraGroups ?? []) {
		const grouped: ChartViewModel = {
			...model,
			type: group.type,
			grouping: group.grouping,
			series: group.series,
			horizontal: false,
		};
		if (group.stock) grouped.stock = group.stock;
		else delete grouped.stock;
		if (group.barGapWidth !== undefined) grouped.barGapWidth = group.barGapWidth;
		else delete grouped.barGapWidth;
		if (group.barOverlap !== undefined) grouped.barOverlap = group.barOverlap;
		else delete grouped.barOverlap;
		const value = group.secondary && frame.secondary ? frame.secondary : frame.value;
		layers.push({ model: grouped, frame: { ...frame, value } });
	}
	clusterLayers(layers)
		.map((layer, order) => ({ layer, order }))
		.sort(
			(a, b) =>
				(RANK[a.layer.model.type] ?? 1) - (RANK[b.layer.model.type] ?? 1) || a.order - b.order,
		)
		.forEach(({ layer }) => paint(layer, out));
	return out.join('');
}
