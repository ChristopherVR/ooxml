import { describe, expect, it } from 'vitest';
import { IOS_CANVAS_MAX_AREA, maxCanvasAreaFor, releaseCanvas } from './canvas-area-cap';
import { isIosWebKit, isWebKitEngine } from './webkit-engine';

const IPHONE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15';
const MAC = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15';
const CHROME = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0';

describe('canvas area cap', () => {
	it('detects iOS and iPadOS WebKit, including the desktop-class iPad user agent', () => {
		expect(isIosWebKit({ userAgent: IPHONE })).toBe(true);
		expect(isIosWebKit({ userAgent: MAC, maxTouchPoints: 5 })).toBe(true);
		expect(isIosWebKit({ userAgent: MAC, maxTouchPoints: 0 })).toBe(false);
		expect(isIosWebKit({ userAgent: CHROME })).toBe(false);
		expect(isIosWebKit(undefined)).toBe(false);
		expect(isWebKitEngine({ userAgent: MAC })).toBe(true);
		expect(isWebKitEngine({ userAgent: CHROME })).toBe(false);
	});
	it('caps the area at 4096 x 4096 on iOS although the per-axis probe reports 16384', () => {
		expect(maxCanvasAreaFor(16384, { userAgent: IPHONE })).toBe(IOS_CANVAS_MAX_AREA);
		expect(maxCanvasAreaFor(16384, { userAgent: CHROME })).toBe(16384 * 16384);
		expect(maxCanvasAreaFor(2048, { userAgent: IPHONE })).toBe(2048 * 2048);
	});
	it('releases a canvas backing store and tolerates a missing canvas', () => {
		const canvas = { width: 4096, height: 4096 };
		releaseCanvas(canvas);
		expect(canvas).toEqual({ width: 0, height: 0 });
		expect(() => releaseCanvas(null)).not.toThrow();
	});
});
