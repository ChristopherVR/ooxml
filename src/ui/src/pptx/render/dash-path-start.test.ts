import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { dashStartPathData } from './dash-path-start';
import { buildStrokeOutline } from './stroke-outline';

const RECT = 'M 0 0 L 100 0 L 100 50 L 0 50 Z';

describe('dashStartPathData', () => {
	it('starts a rectangle at its last vertex, up the closing edge', () => {
		expect(dashStartPathData(RECT, 'rect')).toBe('M 0 50 L 0 0 L 100 0 L 100 50 L 0 50 Z');
	});

	it('starts a triangle at its last vertex', () => {
		expect(dashStartPathData('M 0 50 L 50 0 L 100 50 Z', 'triangle')).toBe(
			'M 100 50 L 0 50 L 50 0 L 100 50 Z',
		);
	});

	it('keeps curves and rotates round the last vertex', () => {
		const rounded = 'M 0 10 A 10 10 0 0 1 10 0 L 90 0 L 90 50 L 10 50 A 10 10 0 0 1 0 40 Z';
		expect(dashStartPathData(rounded, 'roundRect')).toBe(
			'M 0 40 L 0 10 A 10 10 0 0 1 10 0 L 90 0 L 90 50 L 10 50 A 10 10 0 0 1 0 40 Z',
		);
	});

	it('runs an ellipse counter-clockwise from its left-middle point', () => {
		const ellipse =
			'M 0 25 A 50 25 0 0 1 50 0 A 50 25 0 0 1 100 25 A 50 25 0 0 1 50 50 A 50 25 0 0 1 0 25 Z';
		expect(dashStartPathData(ellipse, 'ellipse')).toBe(
			'M 0 25 A 50 25 0 0 0 50 50 A 50 25 0 0 0 100 25 A 50 25 0 0 0 50 0 A 50 25 0 0 0 0 25 Z',
		);
	});

	it('leaves a self-closing figure that is not an ellipse alone', () => {
		const loop = 'M 0 25 C 0 0 50 0 50 25 C 50 50 0 50 0 25 Z';
		expect(dashStartPathData(loop, 'heart')).toBe(loop);
	});

	it('leaves open paths, several sub-paths and unknown commands alone', () => {
		expect(dashStartPathData('M 0 0 L 10 10', 'rect')).toBe('M 0 0 L 10 10');
		expect(dashStartPathData(`${RECT} M 5 5 L 6 6 L 6 7 Z`, 'rect')).toBe(
			`${RECT} M 5 5 L 6 6 L 6 7 Z`,
		);
		expect(dashStartPathData('M 0 0 h 10 v 10 z', 'rect')).toBe('M 0 0 h 10 v 10 z');
	});
});

describe('dashed outline start', () => {
	const rect = (dash: string | undefined): PptxElement =>
		({
			type: 'shape',
			id: 'r',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			shapeType: 'rect',
			shapeStyle: { strokeWidth: 4, strokeColor: '#ffffff', strokeDash: dash },
		}) as PptxElement;

	it('starts a dashed rectangle at its bottom-left corner', () => {
		expect(buildStrokeOutline(rect('dash'))?.d).toBe('M 0 50 L 0 0 L 100 0 L 100 50 L 0 50 Z');
	});

	it('leaves a solid outline in spec order', () => {
		expect(buildStrokeOutline(rect(undefined))?.d).toBe('M 0 0 L 100 0 L 100 50 L 0 50 Z');
	});
});
