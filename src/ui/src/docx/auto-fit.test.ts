import { describe, expect, it } from 'vitest';
import { AUTO_FIT_MIN, autoFitZoom, pinchZoom } from './auto-fit';

describe('autoFitZoom', () => {
	it('keeps 100% when the page fits', () => {
		expect(autoFitZoom(1000, 816)).toBe(100);
		expect(autoFitZoom(816, 816)).toBe(100);
	});
	it('scales a page wider than the canvas to the available width', () => {
		expect(autoFitZoom(359, 816)).toBe(43);
		expect(autoFitZoom(408, 816)).toBe(50);
	});
	it('never goes below the minimum', () => {
		expect(autoFitZoom(50, 816)).toBe(AUTO_FIT_MIN);
	});
	it('ignores an unlaid-out canvas', () => {
		expect(autoFitZoom(0, 816)).toBeUndefined();
		expect(autoFitZoom(300, 0)).toBeUndefined();
	});
});

describe('pinchZoom', () => {
	it('scales and clamps', () => {
		expect(pinchZoom(100, 1.5, 25, 200)).toBe(150);
		expect(pinchZoom(100, 5, 25, 200)).toBe(200);
		expect(pinchZoom(100, 0.1, 25, 200)).toBe(25);
		expect(pinchZoom(80, 0, 25, 200)).toBe(80);
	});
});
