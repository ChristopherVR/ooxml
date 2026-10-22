// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { calculateViewportFit } from './viewport-fit';
import {
	fractionalClientSize,
	measureViewportClientSize,
	observeViewportClientSize,
} from './viewport-size';
import type { ViewportBoxMetrics } from './viewport-size';

/** What a browser reports for an untransformed box laid out at `width` x `height` px. */
function laidOut(width: number, height: number, snapped = Math.round): ViewportBoxMetrics {
	return {
		clientWidth: snapped(width),
		clientHeight: snapped(height),
		offsetWidth: snapped(width),
		offsetHeight: snapped(height),
		rectWidth: width,
		rectHeight: height,
	};
}

const CANVAS = { canvasWidth: 1280, canvasHeight: 720 };

function fittedSize(size: { width: number; height: number }) {
	const { scale } = calculateViewportFit({
		viewportWidth: size.width,
		viewportHeight: size.height,
		...CANVAS,
		fitPadding: 0,
		maxFitScale: null,
	});
	return { width: CANVAS.canvasWidth * scale, height: CANVAS.canvasHeight * scale };
}

describe('fractionalClientSize', () => {
	it('takes the whole-pixel snapping back out of the client size', () => {
		expect(fractionalClientSize(laidOut(801.5, 451.3))).toStrictEqual({
			width: 801.5,
			height: 451.3,
		});
		expect(fractionalClientSize(laidOut(1000.4, 563.6))).toStrictEqual({
			width: 1000.4,
			height: 563.6,
		});
	});

	it('keeps whole-pixel boxes exactly as clientWidth reports them', () => {
		expect(fractionalClientSize(laidOut(960, 540))).toStrictEqual({ width: 960, height: 540 });
	});

	it('subtracts the same snapping from a client box narrowed by a scrollbar or border', () => {
		// 15px scrollbar and a 1px border on each side: client = offset - 17 in whole pixels.
		const metrics = { ...laidOut(801.5, 600), clientWidth: 802 - 17 };
		expect(fractionalClientSize(metrics).width).toBe(801.5 - 17);
	});

	// Sizes laid out at fractional pixels, rounded the way the browser snaps them.
	it.each([
		[801.5, 451.3],
		[801.5, 700],
		[1000.6, 800],
		[1400, 451.5],
		[1200, 563.6],
		[640.75, 500.5],
		[333.33, 999.99],
	])('never lets a slide fitted to %spx x %spx exceed the box', (width, height) => {
		const measured = fittedSize(fractionalClientSize(laidOut(width, height)));
		expect(measured.width).toBeLessThanOrEqual(width + 1e-9);
		expect(measured.height).toBeLessThanOrEqual(height + 1e-9);
		// The fit still fills the limiting axis.
		expect(
			Math.min(Math.abs(width - measured.width), Math.abs(height - measured.height)),
		).toBeLessThan(1e-9);
	});

	it.each([
		[801.5, 700, 'width'],
		[1400, 451.5, 'height'],
	] as const)('the whole-pixel size of %spx x %spx overflows the %s', (width, height, axis) => {
		// What fitting to clientWidth / clientHeight did: half a pixel past the box.
		const naive = fittedSize({ width: Math.round(width), height: Math.round(height) });
		expect(naive[axis] - (axis === 'width' ? width : height)).toBeCloseTo(0.5);
	});

	it('takes fractional borders out of the rect, as browser zoom makes them', () => {
		// 90% browser zoom: a 1px border is 1.11111px, the box is laid out at 801.4931px, and the
		// snapped boxes are 801 and 799. The padding box is 801.4931 - 2.2222 = 799.2709px.
		const zoomed: ViewportBoxMetrics = {
			clientWidth: 799,
			clientHeight: 698,
			offsetWidth: 801,
			offsetHeight: 700,
			rectWidth: 801.493103,
			rectHeight: 699.999,
			borderLeft: 1.11111,
			borderRight: 1.11111,
			borderTop: 1.11111,
			borderBottom: 1.11111,
		};
		const size = fractionalClientSize(zoomed);
		expect(size.width).toBeCloseTo(799.270883, 5);
		expect(size.height).toBeCloseTo(697.776779, 5);
		// Treating the borders as whole pixels carries the rect's error into the client box,
		// 0.22px past the padding box.
		const wholePixelBorders = fractionalClientSize({
			...zoomed,
			borderLeft: undefined,
			borderRight: undefined,
		});
		expect(wholePixelBorders.width - size.width).toBeCloseTo(0.22222, 4);
	});

	it('reserves a pixel for a scrollbar next to fractional borders', () => {
		// 1px borders and a 15px scrollbar: offset 802 - client 785 - borders 2 = 15.
		const metrics: ViewportBoxMetrics = {
			...laidOut(801.5, 600),
			clientWidth: 785,
			borderLeft: 1,
			borderRight: 1,
			borderTop: 1,
			borderBottom: 1,
		};
		expect(fractionalClientSize(metrics).width).toBe(784);
		// Under a pixel between the snapped boxes is rounding, not a scrollbar: 150% zoom on a
		// 1x display snaps a 0.1px border to 0.667px and the boxes to 802 and 800.
		expect(
			fractionalClientSize({
				clientWidth: 800,
				clientHeight: 300,
				offsetWidth: 802,
				offsetHeight: 301,
				rectWidth: 801.5,
				rectHeight: 301.333,
				borderLeft: 0.666667,
				borderRight: 0.666667,
				borderTop: 0.666667,
				borderBottom: 0.666667,
			}).width,
		).toBeCloseTo(800.166667, 5);
	});

	it('reserves a whole pixel when a transform scales the rect', () => {
		// An ancestor scale(0.5): the rect is half the layout size.
		const metrics = { ...laidOut(802, 452), rectWidth: 401, rectHeight: 226 };
		expect(fractionalClientSize(metrics)).toStrictEqual({ width: 801, height: 451 });
	});

	describe('with the computed size (local CSS pixels)', () => {
		/**
		 * Metrics Chromium reports for a pinned viewport (border-box unless noted, no border or
		 * padding unless noted), under real browser zoom and the given ancestor style, and the true
		 * padding box: an `inset: 0` child's rect with the transform removed and the CSS zoom
		 * divided out.
		 */
		const measured = {
			// Plain contexts.
			'90% zoom, 1px borders': {
				metrics: {
					clientWidth: 799,
					clientHeight: 698,
					offsetWidth: 801,
					offsetHeight: 700,
					rectWidth: 801.4931030273438,
					rectHeight: 700.000061,
					borderLeft: 1.11111,
					borderRight: 1.11111,
					borderTop: 1.11111,
					borderBottom: 1.11111,
					cssWidth: 801.493,
					cssHeight: 700,
					boxSizing: 'border-box',
					cssZoom: 1,
				},
				truth: { width: 799.2708740234375, height: 697.77783203125 },
			},
			'90% zoom, content-box with 20px padding and 1px borders': {
				metrics: {
					clientWidth: 800,
					clientHeight: 640,
					offsetWidth: 803,
					offsetHeight: 642,
					rectWidth: 802.7083740234375,
					rectHeight: 642.222229,
					borderLeft: 1.11111,
					borderRight: 1.11111,
					borderTop: 1.11111,
					borderBottom: 1.11111,
					cssWidth: 760.486,
					cssHeight: 600,
					boxSizing: 'content-box',
					paddingLeft: 20,
					paddingRight: 20,
					paddingTop: 20,
					paddingBottom: 20,
					cssZoom: 1,
				},
				truth: { width: 800.486145, height: 640 },
			},
			// The padding is laid out rounded down to 1/64 of a zoomed pixel, so the rect is
			// 0.07px below the computed width plus the computed padding.
			'33% zoom, content-box with 12.3px padding': {
				metrics: {
					clientWidth: 358,
					clientHeight: 476,
					offsetWidth: 358,
					offsetHeight: 476,
					rectWidth: 357.8125,
					rectHeight: 475.99432373046875,
					cssWidth: 333.286,
					cssHeight: 451.468,
					boxSizing: 'content-box',
					paddingLeft: 12.3,
					paddingRight: 12.3,
					paddingTop: 12.3,
					paddingBottom: 12.3,
					cssZoom: 1,
				},
				truth: { width: 357.8125, height: 475.99432373046875 },
			},
			// Transformed or zoomed contexts.
			'ancestor zoom: 1.1 at 90% browser zoom': {
				metrics: {
					clientWidth: 801,
					clientHeight: 700,
					offsetWidth: 801,
					offsetHeight: 700,
					rectWidth: 881.6493530273438,
					rectHeight: 770.000061,
					cssWidth: 801.499,
					cssHeight: 700,
					boxSizing: 'border-box',
					cssZoom: 1.100000023841858,
				},
				truth: { width: 801.4993944710083, height: 700.0000403144134 },
			},
			'ancestor zoom: 0.5 at 50% browser zoom': {
				metrics: {
					clientWidth: 802,
					clientHeight: 700,
					offsetWidth: 802,
					offsetHeight: 700,
					rectWidth: 400.75,
					rectHeight: 350,
					cssWidth: 801.5,
					cssHeight: 700,
					boxSizing: 'border-box',
					cssZoom: 0.5,
				},
				truth: { width: 801.5, height: 700 },
			},
			// 16.5% effective zoom: the snapped client size is 0.33px past the box, and the computed
			// width plus padding 0.17px past it.
			'ancestor zoom: 0.5 at 33% browser zoom, content-box with 12.3px padding': {
				metrics: {
					clientWidth: 358,
					clientHeight: 476,
					offsetWidth: 358,
					offsetHeight: 476,
					rectWidth: 178.83522033691406,
					rectHeight: 237.92613220214844,
					cssWidth: 333.239,
					cssHeight: 451.42,
					boxSizing: 'content-box',
					paddingLeft: 12.3,
					paddingRight: 12.3,
					paddingTop: 12.3,
					paddingBottom: 12.3,
					cssZoom: 0.5,
				},
				truth: { width: 357.6704406738281, height: 475.8522644042969 },
			},
			'ancestor scale(1.0005)': {
				metrics: {
					clientWidth: 1000,
					clientHeight: 800,
					offsetWidth: 1000,
					offsetHeight: 800,
					rectWidth: 1000.8908081054688,
					rectHeight: 800.399963,
					cssWidth: 1000.39,
					cssHeight: 800,
					boxSizing: 'border-box',
					cssZoom: 1,
				},
				truth: { width: 1000.390625, height: 800 },
			},
			'ancestor scale(0.5)': {
				metrics: {
					clientWidth: 802,
					clientHeight: 700,
					offsetWidth: 802,
					offsetHeight: 700,
					rectWidth: 400.75,
					rectHeight: 350,
					cssWidth: 801.5,
					cssHeight: 700,
					boxSizing: 'border-box',
					cssZoom: 1,
				},
				truth: { width: 801.5, height: 700 },
			},
			'ancestor rotate(5deg)': {
				metrics: {
					clientWidth: 802,
					clientHeight: 700,
					offsetWidth: 802,
					offsetHeight: 700,
					rectWidth: 859.4591064453125,
					rectHeight: 767.1915893554688,
					cssWidth: 801.5,
					cssHeight: 700,
					boxSizing: 'border-box',
					cssZoom: 1,
				},
				truth: { width: 801.5, height: 700 },
			},
		} satisfies Record<
			string,
			{ metrics: ViewportBoxMetrics; truth: { width: number; height: number } }
		>;

		/** Zero borders and padding where a case does not list them. */
		function withDefaults(metrics: ViewportBoxMetrics): ViewportBoxMetrics {
			return {
				borderLeft: 0,
				borderRight: 0,
				borderTop: 0,
				borderBottom: 0,
				paddingLeft: 0,
				paddingRight: 0,
				paddingTop: 0,
				paddingBottom: 0,
				...metrics,
			};
		}

		it('is exact in a plain context: 90% zoom, 1px borders', () => {
			const { metrics, truth } = measured['90% zoom, 1px borders'];
			const size = fractionalClientSize(withDefaults(metrics));
			for (const axis of ['width', 'height'] as const) {
				expect(size[axis]).toBeLessThanOrEqual(truth[axis] + 1 / 64);
				expect(truth[axis] - size[axis]).toBeLessThan(0.001);
			}
		});

		it.each([
			'90% zoom, content-box with 20px padding and 1px borders',
			'33% zoom, content-box with 12.3px padding',
		] as const)(
			'stays within the padding rounding below the exact size in a plain context: %s',
			(name) => {
				// Content-box padding is laid out rounded down, so the lower bound of the size is used.
				const { metrics, truth } = measured[name];
				const size = fractionalClientSize(withDefaults(metrics));
				for (const axis of ['width', 'height'] as const) {
					expect(size[axis]).toBeLessThanOrEqual(truth[axis]);
					expect(truth[axis] - size[axis]).toBeLessThan(0.14);
				}
			},
		);

		it('never takes a slightly scaled rect for a padded content box', () => {
			// The padding is laid out 0.07px below its computed size, so scale(1.00003) leaves the
			// rect inside the band a plain rect can be in.
			const { metrics, truth } = measured['33% zoom, content-box with 12.3px padding'];
			const scaled = {
				...metrics,
				rectWidth: metrics.rectWidth * 1.00003,
				rectHeight: metrics.rectHeight * 1.00003,
			};
			const size = fractionalClientSize(withDefaults(scaled));
			expect(size.width).toBeLessThanOrEqual(truth.width);
			expect(size.height).toBeLessThanOrEqual(truth.height);
		});

		it.each([
			'ancestor zoom: 1.1 at 90% browser zoom',
			'ancestor zoom: 0.5 at 50% browser zoom',
			'ancestor zoom: 0.5 at 33% browser zoom, content-box with 12.3px padding',
			'ancestor scale(1.0005)',
			'ancestor scale(0.5)',
			'ancestor rotate(5deg)',
		] as const)('reserves a pixel in a transformed or zoomed context: %s', (name) => {
			// The slide is laid out and painted through the scale there, and its own layout
			// rounding moves its edges by a fraction of a pixel, so an exact size can overflow.
			const { metrics, truth } = measured[name];
			const size = fractionalClientSize(withDefaults(metrics));
			for (const axis of ['width', 'height'] as const) {
				expect(truth[axis] - size[axis]).toBeGreaterThanOrEqual(0.5);
				expect(truth[axis] - size[axis]).toBeLessThanOrEqual(1.5);
			}
		});

		it('treats a CSS zoom as zoomed even when the rect happens to match', () => {
			const { metrics } = measured['90% zoom, 1px borders'];
			const size = fractionalClientSize({ ...metrics, cssZoom: 1.25 });
			// The smaller of the snapped client size and the computed padding box, less a pixel:
			// min(799, 799.27) - 1 and min(698, 697.78) - 1.
			expect(size.width).toBe(798);
			expect(size.height).toBeCloseTo(696.777778, 5);
		});

		it('ignores a missing currentCSSZoom', () => {
			const { metrics, truth } = measured['90% zoom, 1px borders'];
			const size = fractionalClientSize({ ...metrics, cssZoom: undefined });
			expect(size.width).toBeCloseTo(truth.width, 3);
		});

		it('caps the rect of a very wide box to a layout unit above its serialized size', () => {
			// 12000.3px: the computed width (12000.3px) is only good to 0.05px, so a rect up to
			// 0.05px past the box would still look plain.
			const metrics = withDefaults({
				clientWidth: 12000,
				clientHeight: 452,
				offsetWidth: 12000,
				offsetHeight: 452,
				rectWidth: 12000.296875,
				rectHeight: 451.5,
				cssWidth: 12000.3,
				cssHeight: 451.5,
				boxSizing: 'border-box',
				cssZoom: 1,
			});
			expect(fractionalClientSize(metrics).width).toBeCloseTo(12000.25 + 1 / 64, 9);
			// A rect scaled by 1.000004, inside the serialization range, still cannot overflow by
			// more than a layout unit.
			const scaled = fractionalClientSize({ ...metrics, rectWidth: 12000.296875 * 1.000004 });
			expect(scaled.width).toBeLessThanOrEqual(12000.296875 + 1 / 64);
		});

		it('still reserves a pixel for a scrollbar', () => {
			const { metrics } = measured['ancestor scale(1.0005)'];
			const size = fractionalClientSize(withDefaults({ ...metrics, clientWidth: 1000 - 15 }));
			expect(size.width).toBe(984);
		});

		it('falls back to the rect without a usable computed size', () => {
			const { metrics } = measured['90% zoom, 1px borders'];
			for (const fallback of [
				{ cssWidth: Number.NaN },
				{ boxSizing: '' },
				{ boxSizing: 'content-box', paddingLeft: undefined },
			]) {
				const size = fractionalClientSize({ ...withDefaults(metrics), ...fallback });
				expect(size.width).toBeCloseTo(799.270883, 5);
			}
			// And reserves a pixel there for a CSS zoom too.
			const zoomed = fractionalClientSize({
				...withDefaults(measured['ancestor zoom: 1.1 at 90% browser zoom'].metrics),
				cssWidth: Number.NaN,
			});
			expect(zoomed.width).toBe(800);
		});
	});

	it('passes through a box with no layout (detached, display:none, or a DOM without layout)', () => {
		expect(fractionalClientSize(laidOut(0, 0))).toStrictEqual({ width: 0, height: 0 });
		// Test DOMs stub clientWidth alone; offsetWidth and the rect stay 0.
		expect(
			fractionalClientSize({
				clientWidth: 960,
				clientHeight: 540,
				offsetWidth: 0,
				offsetHeight: 0,
				rectWidth: 0,
				rectHeight: 0,
			}),
		).toStrictEqual({ width: 960, height: 540 });
		expect(
			fractionalClientSize({ ...laidOut(10, 10), clientWidth: Number.NaN, rectHeight: Number.NaN }),
		).toStrictEqual({ width: 0, height: 10 });
	});
});

describe('measureViewportClientSize / observeViewportClientSize', () => {
	afterEach(() => {
		vi.unstubAllGlobals();
	});

	function viewport(width: number, height: number): HTMLElement {
		const element = document.createElement('div');
		Object.defineProperties(element, {
			clientWidth: { configurable: true, get: () => Math.round(width) },
			clientHeight: { configurable: true, get: () => Math.round(height) },
			offsetWidth: { configurable: true, get: () => Math.round(width) },
			offsetHeight: { configurable: true, get: () => Math.round(height) },
		});
		element.getBoundingClientRect = () => new DOMRect(0, 0, width, height);
		return element;
	}

	it('reads the element metrics', () => {
		expect(measureViewportClientSize(viewport(801.5, 451.3))).toStrictEqual({
			width: 801.5,
			height: 451.3,
		});
	});

	/** A connected viewport with a computed size (happy-dom resolves inline pixel lengths). */
	function connected(rectWidth: number, cssZoom?: number): HTMLElement {
		const element = viewport(rectWidth, 800);
		Object.defineProperties(element, {
			clientWidth: { configurable: true, get: () => 1000 },
			offsetWidth: { configurable: true, get: () => 1000 },
		});
		if (cssZoom !== undefined) {
			Object.defineProperty(element, 'currentCSSZoom', { configurable: true, value: cssZoom });
		}
		element.style.boxSizing = 'border-box';
		element.style.width = '1000.39px';
		element.style.height = '800px';
		element.style.padding = '0px';
		element.style.borderWidth = '0px';
		element.style.borderStyle = 'solid';
		// Computed styles exist only for a connected element.
		document.body.append(element);
		return element;
	}

	it('uses the rect where it matches the computed size of a connected element', () => {
		const element = connected(1000.390625);
		const { width, height } = measureViewportClientSize(element);
		element.remove();
		expect(width).toBe(1000.390625);
		expect(height).toBe(800);
	});

	it('reserves a pixel where an ancestor transform scales the rect', () => {
		const element = connected(1000.390625 * 1.0005);
		const { width } = measureViewportClientSize(element);
		element.remove();
		expect(width).toBe(999);
	});

	it('reads currentCSSZoom, so an ancestor CSS zoom reserves a pixel', () => {
		const element = connected(1000.390625, 1.1);
		const { width, height } = measureViewportClientSize(element);
		element.remove();
		expect(width).toBe(999);
		expect(height).toBe(799);
	});

	/** A ResizeObserver stub that records what each instance observes. */
	function stubObservers() {
		const observers: { callback: () => void; box: string | undefined; disconnected: boolean }[] =
			[];
		vi.stubGlobal(
			'ResizeObserver',
			class {
				private readonly entry: (typeof observers)[number];
				constructor(callback: () => void) {
					this.entry = { callback, box: undefined, disconnected: false };
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
		return observers;
	}

	function resizable(initial: { width: number; height: number }) {
		const box = { size: initial };
		const element = document.createElement('div');
		Object.defineProperties(element, {
			clientWidth: { get: () => Math.round(box.size.width) },
			clientHeight: { get: () => Math.round(box.size.height) },
			offsetWidth: { get: () => Math.round(box.size.width) },
			offsetHeight: { get: () => Math.round(box.size.height) },
		});
		element.getBoundingClientRect = () => new DOMRect(0, 0, box.size.width, box.size.height);
		return { element, box };
	}

	it('reports synchronously, then on every resize, and stops after dispose', () => {
		const observers = stubObservers();
		const { element, box } = resizable({ width: 801.5, height: 451.3 });
		const reports: { width: number; height: number }[] = [];
		const dispose = observeViewportClientSize(element, (next) => reports.push(next));
		expect(reports).toStrictEqual([{ width: 801.5, height: 451.3 }]);
		box.size = { width: 1000.4, height: 563.7 };
		observers[0]!.callback();
		expect(reports.at(-1)).toStrictEqual({ width: 1000.4, height: 563.7 });
		dispose();
		expect(observers.every((observer) => observer.disconnected)).toBeTruthy();
	});

	it('observes the border box too, so a padding-only change is reported', () => {
		// A content-box viewport with a fixed width: changing its padding changes the client
		// and border boxes, but not the content box a default ResizeObserver watches.
		const observers = stubObservers();
		expect(observers).toHaveLength(0);
		const { element, box } = resizable({ width: 849.5, height: 748 });
		const reports: { width: number; height: number }[] = [];
		observeViewportClientSize(element, (next) => reports.push(next));
		expect(observers.map((observer) => observer.box).toSorted()).toStrictEqual([
			'border-box',
			'content-box',
		]);
		box.size = { width: 801.5, height: 700 };
		observers.find((observer) => observer.box === 'border-box')!.callback();
		expect(reports).toStrictEqual([
			{ width: 849.5, height: 748 },
			{ width: 801.5, height: 700 },
		]);
	});

	it('reports a resize both observers see once', () => {
		const observers = stubObservers();
		const { element, box } = resizable({ width: 801.5, height: 700 });
		const reports: unknown[] = [];
		observeViewportClientSize(element, (next) => reports.push(next));
		// The initial observation each observer delivers is the size already reported.
		for (const observer of observers) {
			observer.callback();
		}
		expect(reports).toHaveLength(1);
		box.size = { width: 640.75, height: 360.4 };
		for (const observer of observers) {
			observer.callback();
		}
		expect(reports).toStrictEqual([
			{ width: 801.5, height: 700 },
			{ width: 640.75, height: 360.4 },
		]);
	});

	it('measures once without ResizeObserver', () => {
		vi.stubGlobal('ResizeObserver', undefined);
		const reports: unknown[] = [];
		const dispose = observeViewportClientSize(viewport(640.75, 360.4), (next) =>
			reports.push(next),
		);
		expect(reports).toStrictEqual([{ width: 640.75, height: 360.4 }]);
		expect(dispose).not.toThrow();
	});
});
