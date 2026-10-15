// Stock charts: a high-low wick per category, with a close tick (high-low-close) or an up/down
// body from open to close (open-high-low-close). The geometry is the shared `../stock-candles`; the
// colours, outlines and bar width are the chart's own `c:hiLowLines` and `c:upDownBars`.
import { STOCK_TICK_WIDTH, stockCandle, stockSlots } from '../stock-candles';
import type { Frame } from './chart-svg-cartesian';
import { esc, line, n, rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';
import { stockView } from './chart-view-stock';

export function stockBars(model: ChartViewModel, frame: Frame, out: string[]): void {
	const slots = stockSlots(model.series.length);
	if (!slots) return;
	// A model built without a resolved style (a hand-made one) paints with Excel's defaults.
	const style = model.stock ?? stockView(undefined, { drawingColor: () => undefined });
	const at = (role: number | undefined, i: number): number | null | undefined =>
		role === undefined ? undefined : model.series[role]?.values[i];
	const closeColor = model.series[slots.close]?.color ?? style.wick.color;
	model.categories.forEach((_, i) => {
		const candle = stockCandle({
			x: frame.band(i),
			band: frame.bandSize,
			high: at(slots.high, i),
			low: at(slots.low, i),
			close: at(slots.close, i),
			open: at(slots.open, i),
			hasOpen: slots.open !== undefined,
			bodyShare: style.bodyShare,
			y: frame.value,
		});
		if (!candle) return;
		out.push(`<g data-chart-series="${slots.close}" data-chart-point="${i}">`);
		out.push(
			line(
				candle.wick.x,
				candle.wick.y1,
				candle.wick.x,
				candle.wick.y2,
				style.wick.color,
				style.wick.width,
			),
		);
		if (candle.body) {
			const { x, y, w, h, up } = candle.body;
			const bar = up ? style.up : style.down;
			const stroke =
				bar.stroke === 'none'
					? ''
					: ` stroke="${esc(bar.stroke)}" stroke-width="${n(bar.strokeWidth)}"`;
			out.push(rect(x, y, w, h, bar.fill).replace('/>', `${stroke}/>`));
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
