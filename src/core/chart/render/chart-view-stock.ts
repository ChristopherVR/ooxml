// A stock group's resolved formatting: the chart's own `c:hiLowLines`, `c:upBars`, `c:downBars` and
// `c:upDownBars/c:gapWidth` where written, Excel's defaults otherwise.
import type { ChartShapeProperties } from '../model-series';
import {
	STOCK_DEFAULT_DOWN_FILL,
	STOCK_DEFAULT_UP_FILL,
	STOCK_DEFAULT_WICK_COLOR,
} from '../stock-candles';
import type { SeriesContext } from './chart-view-series';
import type { ChartSummarySeries, StockSummaryStyle } from './summary';

/** The paint of one up or down bar. */
export interface StockBarView {
	fill: string;
	stroke: string;
	strokeWidth: number;
}

/** A stock group's resolved formatting: wick, up and down bars, and how wide the bars are. */
export interface StockView {
	wick: { color: string; width: number };
	up: StockBarView;
	down: StockBarView;
	/** Share of the category band a bar fills (`c:upDownBars/c:gapWidth`, 150% when unwritten). */
	bodyShare: number;
}

const EMU_PER_PX = 9525;

type Context = Pick<SeriesContext<ChartSummarySeries>, 'drawingColor'>;

function paintOf(fill: ChartShapeProperties['fill'], ctx: Context): string | undefined {
	if (fill?.kind === 'none') return 'none';
	return fill?.kind === 'solid' ? ctx.drawingColor(fill.color) : undefined;
}

function lineOf(
	shape: ChartShapeProperties | undefined,
	ctx: Context,
	color: string,
): { color: string; width: number } {
	const line = shape?.line;
	return {
		color: paintOf(line?.fill, ctx) ?? color,
		width: line?.widthEmu === undefined ? 1 : Math.max(0.5, line.widthEmu / EMU_PER_PX),
	};
}

function barOf(
	shape: ChartShapeProperties | undefined,
	ctx: Context,
	fallback: string,
): StockBarView {
	const stroke = lineOf(shape, ctx, '#404040');
	return {
		fill: paintOf(shape?.fill, ctx) ?? fallback,
		stroke: stroke.color,
		strokeWidth: stroke.width,
	};
}

export function stockView(style: StockSummaryStyle | undefined, ctx: Context): StockView {
	return {
		wick: lineOf(style?.hiLowLines, ctx, STOCK_DEFAULT_WICK_COLOR),
		up: barOf(style?.upBars, ctx, STOCK_DEFAULT_UP_FILL),
		down: barOf(style?.downBars, ctx, STOCK_DEFAULT_DOWN_FILL),
		bodyShare: 1 / (1 + Math.max(0, style?.gapWidth ?? 150) / 100),
	};
}
