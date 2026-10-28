// Resolving a series' trendlines and error bars for the painter: their fit or amount options and
// their stroke. A trendline without its own outline is drawn as Office draws it, a 1.5 pt dotted
// line in the series colour; error bars as a 0.75 pt solid line in the text grey. An outline with
// no fill hides the element.
import type { DrawingColor, DrawingLine } from '../../drawingml/types';
import type { ChartErrorAmountOptions } from '../error-bar-amount';
import type { ChartErrorBarDirection, ChartErrorBarType } from '../model-analysis';
import type { ChartTrendlineFitOptions } from '../trendline-fit';
import { chartPointsToPixels } from './chart-appearance';
import { TEXT_COLOR } from './chart-svg-util';
import type { ChartSummarySeries } from './summary';

/** A resolved stroke: CSS colour, width in pixels and the DrawingML preset dash, when dashed. */
export interface ChartStrokeView {
	color: string;
	width: number;
	dash?: string;
}

export interface ChartTrendlineView {
	fit: ChartTrendlineFitOptions;
	stroke: ChartStrokeView;
}

export interface ChartErrorBarsView {
	direction: ChartErrorBarDirection;
	barType: ChartErrorBarType;
	amount: ChartErrorAmountOptions;
	endCap: boolean;
	stroke: ChartStrokeView;
}

type ColorResolver = (color: DrawingColor | undefined) => string | undefined;

export function strokeOf(
	line: DrawingLine | undefined,
	fallback: ChartStrokeView,
	resolve: ColorResolver,
): ChartStrokeView | undefined {
	if (line?.fill?.kind === 'none') return undefined;
	const color =
		(line?.fill?.kind === 'solid' ? resolve(line.fill.color) : undefined) ?? fallback.color;
	const width =
		line?.widthEmu === undefined ? fallback.width : chartPointsToPixels(line.widthEmu / 12700);
	const dash = line?.dash ?? fallback.dash;
	return { color, width, ...(dash && dash !== 'solid' ? { dash } : {}) };
}

/** The painter's trendlines and error bars of one series, each absent when there are none. */
export function analysisView(
	source: ChartSummarySeries,
	seriesColor: string,
	resolve: ColorResolver,
): { trendlines?: ChartTrendlineView[]; errorBars?: ChartErrorBarsView[] } {
	// A series without a plain paint (no fill, a gradient) gives its trendline the text grey.
	const plain = seriesColor === 'none' || seriesColor.startsWith('url(') ? TEXT_COLOR : seriesColor;
	const trendlines = (source.trendlines ?? []).flatMap((entry): ChartTrendlineView[] => {
		const { name: _name, line, ...fit } = entry;
		const fallback = { color: plain, width: chartPointsToPixels(1.5), dash: 'sysDot' };
		const stroke = strokeOf(line, fallback, resolve);
		return stroke ? [{ fit, stroke }] : [];
	});
	const errorBars = (source.errorBars ?? []).flatMap((entry): ChartErrorBarsView[] => {
		const { direction, barType, noEndCap, line, ...amount } = entry;
		const stroke = strokeOf(line, { color: TEXT_COLOR, width: chartPointsToPixels(0.75) }, resolve);
		return stroke ? [{ direction, barType, amount, endCap: noEndCap !== true, stroke }] : [];
	});
	return {
		...(trendlines.length ? { trendlines } : {}),
		...(errorBars.length ? { errorBars } : {}),
	};
}
