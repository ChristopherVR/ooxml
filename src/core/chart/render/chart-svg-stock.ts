// Stock charts: a high-low wick per category, with a close tick (high-low-close) or an up/down
// body from open to close (open-high-low-close). The geometry is the shared `../stock-candles`.
import {
	STOCK_DEFAULT_DOWN_FILL,
	STOCK_DEFAULT_UP_FILL,
	STOCK_DEFAULT_WICK_COLOR,
	STOCK_TICK_WIDTH,
	stockCandle,
	stockSlots,
} from '../stock-candles';
import type { Frame } from './chart-svg-cartesian';
import { line, n, rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const BODY_STROKE = '#404040';

export function stockBars(model: ChartViewModel, frame: Frame, out: string[]): void {
	const slots = stockSlots(model.series.length);
	if (!slots) return;
	const at = (role: number | undefined, i: number): number | null | undefined =>
		role === undefined ? undefined : model.series[role]?.values[i];
	const closeColor = model.series[slots.close]?.color ?? STOCK_DEFAULT_WICK_COLOR;
	model.categories.forEach((_, i) => {
		const candle = stockCandle({
			x: frame.band(i),
			band: frame.bandSize,
			high: at(slots.high, i),
			low: at(slots.low, i),
			close: at(slots.close, i),
			open: at(slots.open, i),
			hasOpen: slots.open !== undefined,
			y: frame.value,
		});
		if (!candle) return;
		out.push(`<g data-chart-series="${slots.close}" data-chart-point="${i}">`);
		out.push(
			line(candle.wick.x, candle.wick.y1, candle.wick.x, candle.wick.y2, STOCK_DEFAULT_WICK_COLOR),
		);
		if (candle.body) {
			const { x, y, w, h, up } = candle.body;
			out.push(
				rect(x, y, w, h, up ? STOCK_DEFAULT_UP_FILL : STOCK_DEFAULT_DOWN_FILL).replace(
					'/>',
					` stroke="${BODY_STROKE}" stroke-width="${n(1)}"/>`,
				),
			);
		} else if (candle.tick) {
			out.push(
				line(
					candle.tick.x1,
					candle.tick.y,
					candle.tick.x2,
					candle.tick.y,
					closeColor,
					STOCK_TICK_WIDTH,
				),
			);
		}
		out.push('</g>');
	});
}
