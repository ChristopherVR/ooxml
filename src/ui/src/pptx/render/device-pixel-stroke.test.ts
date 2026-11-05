// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import {
	DEVICE_PIXEL_RATIO_VAR,
	DEVICE_PX_VAR,
	deviceStrokeStageStyle,
	neutralizeDeviceStrokes,
	screenBorder,
	screenStrokeWidth,
	snappedStrokeShift,
	snappedStrokeTransform,
	watchDevicePixelRatio,
	withAuthoredStrokeWidths,
} from './device-pixel-stroke';
import { prepareExportClone } from './export-clone';

describe('screenStrokeWidth', () => {
	it('holds a painted width at one device pixel or more, authored where the var is unset', () => {
		expect(screenStrokeWidth(0.13)).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(0.13px, max(var(--pptx-device-px, 0px), 0.001px)))',
		);
		expect(screenStrokeWidth(2)).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(2px, max(var(--pptx-device-px, 0px), 0.001px)))',
		);
	});

	it('keeps a zero, negative, missing or non-finite width at 0px (no invented hairline)', () => {
		for (const width of [0, -1, undefined, Number.NaN, Number.POSITIVE_INFINITY]) {
			expect(screenStrokeWidth(width)).toBe('0px');
		}
	});

	it('never emits exponent notation, which CSS rejects', () => {
		expect(screenStrokeWidth(1e-7)).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(0.0001px, max(var(--pptx-device-px, 0px), 0.001px)))',
		);
		expect(screenStrokeWidth(1 / 3)).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(0.3333px, max(var(--pptx-device-px, 0px), 0.001px)))',
		);
	});

	it('builds a border shorthand from the same width', () => {
		expect(screenBorder(0.5, 'dashed', '#000')).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(0.5px, max(var(--pptx-device-px, 0px), 0.001px))) dashed #000',
		);
		expect(screenBorder(1, undefined, 'red')).toBe(
			'max(max(var(--pptx-device-px, 0px), 0.001px), round(1px, max(var(--pptx-device-px, 0px), 0.001px))) solid red',
		);
		expect(screenBorder(0, 'solid', 'red')).toBe('0px solid red');
	});
});

describe('deviceStrokeStageStyle', () => {
	it('publishes one device pixel in unscaled slide px, against the live ratio', () => {
		expect(deviceStrokeStageStyle(0.5)).toEqual({
			[DEVICE_PX_VAR]: `calc(1px / (0.5 * var(${DEVICE_PIXEL_RATIO_VAR}, 1)))`,
		});
	});

	it('publishes nothing for a degenerate scale, so authored widths stay', () => {
		expect(deviceStrokeStageStyle(0)).toEqual({});
		expect(deviceStrokeStageStyle(Number.NaN)).toEqual({});
	});
});

describe('watchDevicePixelRatio', () => {
	function fakeWindow(ratio: number) {
		const props = new Map<string, string>();
		const listeners: Array<() => void> = [];
		const queries: string[] = [];
		const win = {
			devicePixelRatio: ratio,
			document: {
				documentElement: { style: { setProperty: (k: string, v: string) => props.set(k, v) } },
			},
			matchMedia: (query: string) => {
				queries.push(query);
				return {
					addEventListener: (_type: 'change', listener: () => void) => listeners.push(listener),
				};
			},
		};
		return { win, props, listeners, queries };
	}

	it('mirrors the ratio onto the root and re-arms on every change', () => {
		const { win, props, listeners, queries } = fakeWindow(2);
		watchDevicePixelRatio(win);
		expect(props.get(DEVICE_PIXEL_RATIO_VAR)).toBe('2');
		expect(queries).toEqual(['(resolution: 2dppx)']);

		win.devicePixelRatio = 1.25;
		listeners.shift()?.();
		expect(props.get(DEVICE_PIXEL_RATIO_VAR)).toBe('1.25');
		expect(queries.at(-1)).toBe('(resolution: 1.25dppx)');
	});

	it('installs once per window and falls back to 1 for a bogus ratio', () => {
		const { win, props, queries } = fakeWindow(0);
		watchDevicePixelRatio(win);
		watchDevicePixelRatio(win);
		expect(props.get(DEVICE_PIXEL_RATIO_VAR)).toBe('1');
		expect(queries).toHaveLength(1);
	});
});

describe('export keeps authored widths', () => {
	function stageWithStroke(): { host: HTMLElement; stage: HTMLElement; path: SVGPathElement } {
		const host = document.createElement('div');
		const stage = document.createElement('div');
		stage.style.setProperty(DEVICE_PX_VAR, 'calc(1px / (0.5 * var(--pptx-dpr, 1)))');
		const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
		path.style.setProperty('stroke-width', screenStrokeWidth(0.13));
		stage.appendChild(path);
		host.appendChild(stage);
		return { host, stage, path };
	}

	it('pins the device pixel to 0px on every nested stage, leaving stage-free nodes alone', () => {
		const { host, stage, path } = stageWithStroke();
		neutralizeDeviceStrokes(host);
		expect(stage.style.getPropertyValue(DEVICE_PX_VAR)).toBe('0px');
		expect(host.getAttribute('style')).toBeNull();
		expect(path.style.getPropertyValue(DEVICE_PX_VAR)).toBe('');
	});

	it('pins a captured element that sits INSIDE a stage (html2canvas keeps its ancestors)', () => {
		const { stage } = stageWithStroke();
		const shape = document.createElement('div');
		stage.appendChild(shape);
		neutralizeDeviceStrokes(shape);
		expect(shape.style.getPropertyValue(DEVICE_PX_VAR)).toBe('0px');
	});

	it('runs as part of the shared export clone pass', () => {
		const { host, stage } = stageWithStroke();
		prepareExportClone(host);
		expect(stage.style.getPropertyValue(DEVICE_PX_VAR)).toBe('0px');
	});

	it('restores the live values after a computed-style read', () => {
		const { host, stage, path } = stageWithStroke();
		const seen = withAuthoredStrokeWidths(host, () => stage.style.getPropertyValue(DEVICE_PX_VAR));
		expect(seen).toBe('0px');
		expect(stage.style.getPropertyValue(DEVICE_PX_VAR)).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);
		expect(host.getAttribute('style')).toBeNull();
		expect(path.style.getPropertyValue(DEVICE_PX_VAR)).toBe('');
	});

	it('restores only the variable, never re-parsing the style attribute', () => {
		// A browser serialises inline styles with rounded numbers, so writing the
		// attribute back would shift a framework-set transform on the live stage.
		const { host, stage } = stageWithStroke();
		stage.style.setProperty('transform', 'scale(0.9958333333)');
		const writes: string[] = [];
		const setAttribute = stage.setAttribute.bind(stage);
		stage.setAttribute = (name: string, value: string) => {
			writes.push(name);
			setAttribute(name, value);
		};
		withAuthoredStrokeWidths(host, () => undefined);
		expect(writes).toEqual([]);
		expect(stage.style.getPropertyValue('transform')).toBe('scale(0.9958333333)');
		expect(stage.style.getPropertyValue(DEVICE_PX_VAR)).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);
	});

	it('leaves no style attribute on a captured element that had none', () => {
		const { stage } = stageWithStroke();
		const shape = document.createElement('div');
		stage.appendChild(shape);
		withAuthoredStrokeWidths(shape, () => undefined);
		expect(shape.getAttribute('style')).toBeNull();
		expect(stage.style.getPropertyValue(DEVICE_PX_VAR)).toBe(
			'calc(1px / (0.5 * var(--pptx-dpr, 1)))',
		);
	});
});

describe('PowerPoint thin-line rule (measured on PowerPoint 16 exports)', () => {
	const STEP = 'max(var(--pptx-device-px, 0px), 0.001px)';

	it('quantizes a width to whole device pixels instead of keeping the authored fraction', () => {
		expect(screenStrokeWidth(1.3333)).toBe(`max(${STEP}, round(1.3333px, ${STEP}))`);
		expect(screenStrokeWidth(1.3333)).not.toBe('max(1.3333px, var(--pptx-device-px, 0px))');
	});

	it('floors a width off screen only when asked, and never inside a stage', () => {
		expect(screenStrokeWidth(0.5)).not.toContain('clamp(');
		const floored = screenStrokeWidth(0.5, 1);
		expect(floored).toContain('clamp(0px, calc(var(--pptx-device-px, 1000px) - 100px), 1px)');
		expect(floored).toContain(`round(0.5px, ${STEP})`);
	});

	it('snaps an axis-aligned stroke by round(centre) - floor(n / 2) device pixels', () => {
		const shift = snappedStrokeShift(27, 0.3333);
		expect(shift).toContain(`round(27px, ${STEP})`);
		expect(shift).toContain('round(down, calc(');
		expect(shift).toMatch(/- 27px\)$/u);
	});

	it('shifts a horizontal line in y only, a vertical one in x only, a box in both', () => {
		expect(snappedStrokeTransform('h', 10, 20, 1)).toMatch(/^translate\(0px, calc\(round\(20px,/u);
		expect(snappedStrokeTransform('v', 10, 20, 1)).toMatch(
			/^translate\(calc\(round\(10px,.*\), 0px\)$/u,
		);
		const box = snappedStrokeTransform('box', 10, 20, 1) ?? '';
		expect(box).toContain('calc(round(10px,');
		expect(box).toContain('calc(round(20px,');
	});

	it('shifts nothing for a missing, zero or non-finite width or position', () => {
		expect(snappedStrokeTransform('h', 1, 1, 0)).toBeUndefined();
		expect(snappedStrokeTransform('h', 1, 1, undefined)).toBeUndefined();
		expect(snappedStrokeShift(Number.NaN, 1)).toBe('0px');
	});
});
