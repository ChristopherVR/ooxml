import { describe, expect, it } from 'vitest';

import { stockCandle, stockSlots } from './stock-candles';

const y = (v: number): number => 100 - v;

describe('stockSlots', () => {
	it('reads three series as HLC and four as OHLC', () => {
		expect(stockSlots(3)).toEqual({ high: 0, low: 1, close: 2 });
		expect(stockSlots(4)).toEqual({ open: 0, high: 1, low: 2, close: 3 });
	});

	it('can take the last series as close and rejects fewer than three', () => {
		expect(stockSlots(5, true)?.close).toBe(4);
		expect(stockSlots(2)).toBeUndefined();
	});
});

describe('stockCandle', () => {
	it('draws a body from open to close for OHLC and flags up and down', () => {
		const up = stockCandle({
			x: 50,
			band: 40,
			high: 30,
			low: 10,
			open: 15,
			close: 25,
			hasOpen: true,
			y,
		});
		expect(up?.wick).toEqual({ x: 50, y1: 70, y2: 90 });
		expect(up?.body).toEqual({ x: 40, y: 75, w: 20, h: 10, up: true });
		const down = stockCandle({
			x: 50,
			band: 40,
			high: 30,
			low: 10,
			open: 25,
			close: 15,
			hasOpen: true,
			y,
		});
		expect(down?.body?.up).toBe(false);
	});

	it('gives HLC a close tick and no body', () => {
		const candle = stockCandle({
			x: 50,
			band: 40,
			high: 30,
			low: 10,
			close: 20,
			hasOpen: false,
			y,
		});
		expect(candle?.body).toBeUndefined();
		expect(candle?.tick).toEqual({ x1: 50, x2: 58, y: 80 });
	});

	it('skips a category with no high or low and keeps a minimum body height', () => {
		expect(
			stockCandle({ x: 0, band: 10, high: null, low: 1, close: 1, hasOpen: false, y }),
		).toBeUndefined();
		const flat = stockCandle({
			x: 0,
			band: 10,
			high: 5,
			low: 1,
			open: 3,
			close: 3,
			hasOpen: true,
			y,
		});
		expect(flat?.body?.h).toBe(1);
	});
});
