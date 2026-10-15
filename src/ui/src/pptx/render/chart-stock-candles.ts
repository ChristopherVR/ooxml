/**
 * chart-stock-candles.ts: shared HLC/OHLC candle geometry (hi-lo wick, close
 * tick, candle body) for a stock chart's price series.
 *
 * PowerPoint draws two distinct shapes depending on whether an Open series is
 * present:
 *   - HLC (High/Low/Close, no Open): a vertical hi-lo wick with NO body, plus
 *     a short horizontal tick at the close value (`c:upDownBars` is absent).
 *   - OHLC (Open/High/Low/Close): the same wick, plus a candle body from open
 *     to close, filled from `c:upDownBars/c:upBars`(`downBars`) - never a
 *     hardcoded colour, and never green/red by default (PowerPoint's own
 *     default is white-ish up / dark down, see `DEFAULT_UP_FILL`/`DEFAULT_DOWN_FILL`).
 *
 * Used by both a standalone stock chart (`chart-combo-stock.ts`) and the
 * price portion of a volume+stock combo (`chart-combo.ts`, where a
 * `c:barChart` volume series sits alongside `c:stockChart` price series on a
 * secondary axis): one geometry function fixes the visual for both call
 * sites at once (CLAUDE.md Rule 2).
 *
 * @module chart-stock-candles
 */
import {
	STOCK_DEFAULT_DOWN_FILL,
	STOCK_DEFAULT_UP_FILL,
	STOCK_DEFAULT_WICK_COLOR,
	STOCK_TICK_WIDTH,
	stockCandle,
} from 'ooxml-core/chart';
import type { PptxChartData, PptxChartSeries } from 'ooxml-core/pptx';

import { seriesColor } from './chart-helpers';
import { buildStockCloseLabel } from './chart-stock-close-label';
import type {
	PlotLayout,
	SvgLine,
	SvgPrimitive,
	SvgRect,
	SvgText,
	ValueRange,
} from './chart-view-model';
import { valueToY } from './chart-view-model';

/** The resolved price series for one candle group, in OOXML stock-series order. */
export interface StockCandleSeries {
	/** Present only for OHLC; its absence is what selects the HLC (tick) rendering. */
	open?: PptxChartSeries;
	high: PptxChartSeries;
	low: PptxChartSeries;
	close: PptxChartSeries;
	/** `close`'s index within `chartData.series`, for its own colour and data-point tagging. */
	closeIndex: number;
}

/**
 * Build the wick + (tick or body) primitives for every category. `range` and
 * `xPositions` are the caller's own axis placement, so this same function
 * serves a standalone stock chart (primary axis) and a volume-combo's price
 * series (typically the secondary axis) alike.
 */
export function computeStockCandlePrimitives(
	seriesSet: StockCandleSeries,
	chartData: PptxChartData,
	layout: PlotLayout,
	range: ValueRange,
	catCount: number,
	sourceIndices: ReadonlyArray<number>,
	xPositions: ReadonlyArray<number> | undefined,
): SvgPrimitive[] {
	const { open, high, low, close, closeIndex } = seriesSet;
	const primitives: SvgPrimitive[] = [];
	const barGroupWidth = layout.plotWidth / Math.max(catCount, 1);
	const wickColor = chartData.hiLowLines?.color ?? STOCK_DEFAULT_WICK_COLOR;
	const wickWidth = chartData.hiLowLines?.width ?? 1;
	const closeColor = seriesColor(
		close,
		closeIndex,
		chartData.style?.styleId,
		chartData.colorPalette,
	);
	const upDownBars = chartData.upDownBars;

	for (let displayIndex = 0; displayIndex < catCount; displayIndex++) {
		const sourceIndex = sourceIndices[displayIndex] ?? displayIndex;
		const candle = stockCandle({
			x:
				xPositions?.[displayIndex] ??
				layout.plotLeft + barGroupWidth * displayIndex + barGroupWidth / 2,
			band: barGroupWidth,
			high: high.values[sourceIndex],
			low: low.values[sourceIndex],
			close: close.values[sourceIndex],
			open: open?.values[sourceIndex],
			hasOpen: open !== undefined,
			y: (value) => valueToY(value, range, layout.plotTop, layout.plotBottom),
		});
		if (!candle) {
			continue;
		}
		primitives.push({
			kind: 'line',
			x1: candle.wick.x,
			y1: candle.wick.y1,
			x2: candle.wick.x,
			y2: candle.wick.y2,
			stroke: wickColor,
			strokeWidth: wickWidth,
		} satisfies SvgLine);
		if (candle.body) {
			// OHLC: a filled candle body from open to close, coloured by the
			// chart's own up/down bar formatting (never a hardcoded palette).
			const barProps = candle.body.up ? upDownBars?.upBars : upDownBars?.downBars;
			primitives.push({
				kind: 'rect',
				x: candle.body.x,
				y: candle.body.y,
				w: candle.body.w,
				h: candle.body.h,
				fill:
					barProps?.fillColor ?? (candle.body.up ? STOCK_DEFAULT_UP_FILL : STOCK_DEFAULT_DOWN_FILL),
				part: { role: 'dataPoint', seriesIndex: closeIndex, pointIndex: sourceIndex },
			} satisfies SvgRect);
		} else if (candle.tick) {
			// HLC: no body. A short tick to the right of the wick marks the close,
			// coloured like the Close series itself (matches PowerPoint's legend).
			primitives.push({
				kind: 'line',
				x1: candle.tick.x1,
				y1: candle.tick.y,
				x2: candle.tick.x2,
				y2: candle.tick.y,
				stroke: closeColor,
				strokeWidth: STOCK_TICK_WIDTH,
			} satisfies SvgLine);
		}
	}
	return primitives;
}

/** One `chartData.series` entry tagged `seriesChartType: 'stock'` (combo charts only). */
export interface ComboStockSeriesEntry {
	series: PptxChartSeries;
	index: number;
}

/**
 * Pick out a combo chart's stock-tagged series (a volume+stock combo writes
 * `c:barChart` for the volume bar and `c:stockChart` for the HLC/OHLC price
 * series; the loader tags the latter with `seriesChartType: 'stock'`), in
 * their original `chartData.series` order.
 */
export function findComboStockSeries(
	series: ReadonlyArray<PptxChartSeries>,
): ComboStockSeriesEntry[] {
	return series
		.map((s, index) => ({ series: s, index }))
		.filter((entry) => entry.series.seriesChartType === 'stock');
}

/**
 * Build the candle geometry AND close-value data labels for a combo chart's
 * stock-tagged series (the volume bar is rendered separately by the caller).
 * Mirrors `buildStockViewModel`'s own candle + label loop so a volume+stock
 * combo matches a standalone stock chart pixel-for-pixel modulo axis choice.
 */
export function computeComboStockOverlay(
	stockEntries: ReadonlyArray<ComboStockSeriesEntry>,
	chartData: PptxChartData,
	layout: PlotLayout,
	range: ValueRange,
	catCount: number,
	sourceIndices: ReadonlyArray<number>,
	xPositions: ReadonlyArray<number> | undefined,
): { primitives: SvgPrimitive[]; dataLabels: SvgText[] } {
	if (stockEntries.length < 3) {
		return { primitives: [], dataLabels: [] };
	}
	const hasFour = stockEntries.length >= 4;
	const open = hasFour ? stockEntries[0].series : undefined;
	const high = stockEntries[hasFour ? 1 : 0].series;
	const low = stockEntries[hasFour ? 2 : 1].series;
	const closeEntry = stockEntries[stockEntries.length - 1];

	const primitives = computeStockCandlePrimitives(
		{ open, high, low, close: closeEntry.series, closeIndex: closeEntry.index },
		chartData,
		layout,
		range,
		catCount,
		sourceIndices,
		xPositions,
	);

	const dataLabels: SvgText[] = [];
	if (chartData.style?.hasDataLabels) {
		const barGroupWidth = layout.plotWidth / Math.max(catCount, 1);
		for (let displayIndex = 0; displayIndex < catCount; displayIndex++) {
			const sourceIndex = sourceIndices[displayIndex] ?? displayIndex;
			const close = closeEntry.series.values[sourceIndex];
			if (close === undefined) {
				continue;
			}
			const cx =
				xPositions?.[displayIndex] ??
				layout.plotLeft + barGroupWidth * displayIndex + barGroupWidth / 2;
			const closeY = valueToY(close, range, layout.plotTop, layout.plotBottom);
			const label = buildStockCloseLabel(
				chartData,
				closeEntry.series,
				sourceIndex,
				close,
				cx,
				closeY,
				{
					width: layout.svgWidth,
					height: layout.svgHeight,
				},
			);
			if (label) {
				dataLabels.push(label);
			}
		}
	}
	return { primitives, dataLabels };
}
