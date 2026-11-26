import { describe, expect, it } from 'vitest';

import { measureExportRoot, uniformZoomOf, untransformExportClone } from './export-root-size';

/** An element whose layout and on-screen box are given, with a computed `transform`. */
function element(layout: [number, number], box: [number, number], transform: string): HTMLElement {
	const view = { getComputedStyle: () => ({ transform }) };
	return {
		offsetWidth: layout[0],
		offsetHeight: layout[1],
		getBoundingClientRect: () => ({ width: box[0], height: box[1] }),
		ownerDocument: { defaultView: view },
	} as unknown as HTMLElement;
}

describe('uniformZoomOf', () => {
	it('reads a uniform positive scale, translated or not', () => {
		expect(uniformZoomOf('matrix(0.546875, 0, 0, 0.546875, 0, 0)')).toBe(0.546875);
		expect(uniformZoomOf('matrix(2, 0, 0, 2, 10, 20)')).toBe(2);
	});

	it('refuses rotation, flips, skew, anisotropic scale, 3D and none', () => {
		expect(uniformZoomOf('matrix(0.7071, 0.7071, -0.7071, 0.7071, 0, 0)')).toBeUndefined();
		expect(uniformZoomOf('matrix(-1, 0, 0, 1, 0, 0)')).toBeUndefined();
		expect(uniformZoomOf('matrix(1, 0, 0.3, 1, 0, 0)')).toBeUndefined();
		expect(uniformZoomOf('matrix(1, 0, 0, 0.5, 0, 0)')).toBeUndefined();
		expect(
			uniformZoomOf('matrix3d(1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1)'),
		).toBeUndefined();
		expect(uniformZoomOf('none')).toBeUndefined();
	});
});

describe('measureExportRoot', () => {
	it('exports a zoomed stage at its layout size without the zoom, whatever the window', () => {
		for (const zoom of [0.546875, 0.871875]) {
			const stage = element(
				[1280, 720],
				[1280 * zoom, 720 * zoom],
				`matrix(${zoom}, 0, 0, ${zoom}, 0, 0)`,
			);
			expect(measureExportRoot(stage)).toEqual({
				width: 1280,
				height: 720,
				untransformedRoot: true,
			});
		}
	});

	it('keeps a rotated element on its on-screen box and transform', () => {
		const shape = element([100, 50], [106, 106], 'matrix(0.7071, 0.7071, -0.7071, 0.7071, 0, 0)');
		expect(measureExportRoot(shape)).toEqual({ width: 106, height: 106, untransformedRoot: false });
	});

	it('falls back to the layout size when the box is empty', () => {
		expect(measureExportRoot(element([300, 200], [0, 0], 'none'))).toEqual({
			width: 300,
			height: 200,
			untransformedRoot: false,
		});
	});
});

it('untransformExportClone clears every transform property', () => {
	const set = new Map<string, string>();
	const clone = { style: { setProperty: (name: string, value: string) => set.set(name, value) } };
	untransformExportClone(clone as unknown as HTMLElement);
	expect(Object.fromEntries(set)).toEqual({
		transform: 'none',
		translate: 'none',
		rotate: 'none',
		scale: 'none',
	});
});
