/**
 * Stock-chart (high-low-close and open-high-low-close) geometry shared by every format. Moved from
 * the PowerPoint renderer (`chart-stock-candles.ts`) so the shared SVG painter draws the same
 * candles.
 *
 * Office draws two shapes depending on whether an Open series is present:
 * - HLC (no Open): a vertical high-low wick with no body, plus a short horizontal tick at the
 *   close value.
 * - OHLC: the same wick, plus a candle body from open to close, filled from `c:upDownBars`
 *   (`c:upBars` / `c:downBars`); Office's own default is a white up bar and a dark down bar.
 */

/** Office's default up/down bar fill when a stock chart carries no `c:upDownBars` formatting. */
export const STOCK_DEFAULT_UP_FILL = '#FFFFFF';
export const STOCK_DEFAULT_DOWN_FILL = '#404040';
export const STOCK_DEFAULT_WICK_COLOR = '#334155';
/** Stroke width of the HLC close tick. */
export const STOCK_TICK_WIDTH = 1.5;

/** Which series of a stock group plays which price role (indices into the group's series). */
export interface StockSlots {
	/** Present only for OHLC; its absence selects the HLC (tick) rendering. */
	open?: number;
	high: number;
	low: number;
	close: number;
}

/**
 * The roles of `count` series in OOXML stock-series order: three are High/Low/Close, four or more
 * are Open/High/Low/Close. `closeLast` takes the final series as Close (a volume combination
 * tags a stock group's series individually), otherwise Close is the third or fourth one.
 */
export function stockSlots(count: number, closeLast = false): StockSlots | undefined {
	if (count < 3) return undefined;
	const four = count >= 4;
	return {
		...(four ? { open: 0 } : {}),
		high: four ? 1 : 0,
		low: four ? 2 : 1,
		close: closeLast ? count - 1 : four ? 3 : 2,
	};
}

type Value = number | null | undefined;

export interface StockCandleInput {
	/** Centre of the category band. */
	x: number;
	/** Width of the category band. */
	band: number;
	high: Value;
	low: Value;
	close: Value;
	/** Present for OHLC. */
	open?: Value;
	hasOpen: boolean;
	/** Share of the band a body fills (default 0.5, a gap width of 100). */
	bodyShare?: number;
	/** Value to pixel. */
	y: (value: number) => number;
}

/** The marks of one category: a wick, plus a body (OHLC) or a close tick (HLC). */
export interface StockCandle {
	wick: { x: number; y1: number; y2: number };
	body?: { x: number; y: number; w: number; h: number; up: boolean };
	tick?: { x1: number; x2: number; y: number };
}

const present = (v: Value): v is number => v !== null && v !== undefined;

/** The candle of one category, or undefined when its high or low is missing. */
export function stockCandle(input: StockCandleInput): StockCandle | undefined {
	const { high, low, close, open, x, band, y } = input;
	if (!present(high) || !present(low)) return undefined;
	const candleWidth = band * (input.bodyShare ?? 0.5);
	const candle: StockCandle = { wick: { x, y1: y(high), y2: y(low) } };
	if (input.hasOpen && present(open) && present(close)) {
		const openY = y(open);
		const closeY = y(close);
		candle.body = {
			x: x - candleWidth / 2,
			y: Math.min(openY, closeY),
			w: candleWidth,
			h: Math.max(Math.abs(openY - closeY), 1),
			up: close >= open,
		};
	} else if (present(close)) {
		candle.tick = {
			x1: x,
			x2: x + Math.max(candleWidth * 0.4, 4),
			y: y(close),
		};
	}
	return candle;
}
