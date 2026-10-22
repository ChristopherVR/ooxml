/**
 * slide-canvas-viewport.test.ts: Unit tests for the viewport-background press
 * guard. Ports the React `useCanvasEventHandlers` coverage: a direct hit on the
 * scrollable viewport background deselects, while bubbled child events do not.
 */

import { afterEach, describe, expect, it, vi } from 'vitest';

import { CanvasFitService } from './canvas-fit.service';
import { isViewportBackgroundPressTarget } from './slide-canvas-helpers';

describe('isViewportBackgroundPressTarget', () => {
	it('returns true when the press lands directly on the viewport background', () => {
		const viewport = {} as EventTarget;
		expect(isViewportBackgroundPressTarget(viewport, viewport)).toBeTruthy();
	});

	it('returns false for bubbled child events (target differs from currentTarget)', () => {
		const viewport = {} as EventTarget;
		const child = {} as EventTarget;
		expect(isViewportBackgroundPressTarget(child, viewport)).toBeFalsy();
	});

	it('returns false when target differs from a null currentTarget', () => {
		const viewport = {} as EventTarget;
		expect(isViewportBackgroundPressTarget(null, viewport)).toBeFalsy();
		expect(isViewportBackgroundPressTarget(viewport, null)).toBeFalsy();
	});
});

describe('canvas viewport fit policy', () => {
	afterEach(() => {
		vi.unstubAllGlobals();
	});

	function setup(width = 960, height = 540) {
		const viewport = document.createElement('div');
		Object.defineProperties(viewport, {
			clientWidth: { get: () => width },
			clientHeight: { get: () => height },
		});
		const state: {
			autoFit: boolean;
			fitPadding?: number | { horizontal: number; vertical: number };
			maxFitScale?: number | null;
		} = { autoFit: true };
		const service = new CanvasFitService();
		service.bind({
			autoFit: () => state.autoFit,
			viewportElement: () => viewport,
			canvasSize: () => ({ width: 960, height: 540 }),
			fitOptions: () => state,
		});
		return { service, state };
	}

	it.each([
		[801.5, 700],
		[1200, 451.5],
	])('fits a %spx x %spx viewport without exceeding it', (width, height) => {
		// What a browser reports for a box laid out at a fractional size: whole-pixel client and
		// offset sizes, and the unrounded rect.
		const viewport = document.createElement('div');
		Object.defineProperties(viewport, {
			clientWidth: { get: () => Math.round(width) },
			clientHeight: { get: () => Math.round(height) },
			offsetWidth: { get: () => Math.round(width) },
			offsetHeight: { get: () => Math.round(height) },
		});
		viewport.getBoundingClientRect = () => new DOMRect(0, 0, width, height);
		const service = new CanvasFitService();
		service.bind({
			autoFit: () => true,
			viewportElement: () => viewport,
			canvasSize: () => ({ width: 960, height: 540 }),
			fitOptions: () => ({ fitPadding: 0, maxFitScale: null }),
		});
		service.recompute();
		const fit = service.fitScale();
		expect(960 * fit).toBeLessThanOrEqual(width + 1e-9);
		expect(540 * fit).toBeLessThanOrEqual(height + 1e-9);
		expect(fit).toBeCloseTo(Math.min(width / 960, height / 540), 9);
	});

	it('observes both boxes, so a padding-only change refits', () => {
		const observers: { box: string | undefined; callback: () => void; disconnected: boolean }[] =
			[];
		vi.stubGlobal(
			'ResizeObserver',
			class {
				private readonly entry: (typeof observers)[number];
				constructor(callback: () => void) {
					this.entry = { box: undefined, callback, disconnected: false };
					observers.push(this.entry);
				}
				observe(_node: Element, options?: ResizeObserverOptions) {
					this.entry.box = options?.box;
				}
				disconnect() {
					this.entry.disconnected = true;
				}
			},
		);
		// A content-box viewport losing its padding: the client size drops, the content box
		// a default ResizeObserver watches stays put.
		let size = { width: 1008, height: 588 };
		const viewport = document.createElement('div');
		Object.defineProperties(viewport, {
			clientWidth: { get: () => size.width },
			clientHeight: { get: () => size.height },
		});
		const service = new CanvasFitService();
		service.bind({
			autoFit: () => true,
			viewportElement: () => viewport,
			canvasSize: () => ({ width: 960, height: 540 }),
			fitOptions: () => ({ fitPadding: 0, maxFitScale: null }),
		});
		const stop = service.observe(viewport);
		expect(service.fitScale()).toBe(1.05);
		size = { width: 960, height: 540 };
		observers.find((observer) => observer.box === 'border-box')!.callback();
		expect(service.fitScale()).toBe(1);
		stop();
		expect(observers.every((observer) => observer.disconnected)).toBeTruthy();
	});

	it('preserves default padding and supports live padding updates', () => {
		const { service, state } = setup();
		service.recompute();
		expect(service.fitScale()).toBeCloseTo(508 / 540);
		state.fitPadding = 0;
		service.recompute();
		expect(service.fitScale()).toBe(1);
		state.fitPadding = { horizontal: 80, vertical: 0 };
		service.recompute();
		expect(service.fitScale()).toBeCloseTo(800 / 960);
	});

	it('supports uncapped enlargement, finite caps, and the autoFit bypass', () => {
		const { service, state } = setup(1920, 1080);
		service.recompute();
		expect(service.fitScale()).toBe(1);
		state.fitPadding = 0;
		state.maxFitScale = null;
		service.recompute();
		expect(service.fitScale()).toBe(2);
		state.maxFitScale = 1.5;
		service.recompute();
		expect(service.fitScale()).toBe(1.5);
		state.autoFit = false;
		service.recompute();
		expect(service.fitScale()).toBe(1);
	});

	it('reserves the same horizontal allowance as the other editor bindings', () => {
		const { service } = setup(968, 1080);
		service.recompute();
		expect(service.fitScale()).toBe(1);
	});

	it('insets the slide by the phone margin below the mobile breakpoint', () => {
		const { service } = setup(488, 540);
		service.recompute();
		expect(service.fitScale()).toBeCloseTo((488 - 24) / 960);
	});

	it('falls back safely for invalid options and unavailable measurements', () => {
		const { service, state } = setup();
		state.fitPadding = -10;
		state.maxFitScale = Number.NaN;
		service.recompute();
		expect(service.fitScale()).toBeCloseTo(508 / 540);
		const unavailable = setup(0, 0).service;
		unavailable.recompute();
		expect(unavailable.fitScale()).toBe(1);
	});
});
