import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	getSplitStrokePresetPaths,
	getStrokeOnlyPresetPaths,
	isStrokeOnlyPresetElement,
	splitStrokePresetPathData,
	strokeOnlyPresetPathData,
} from './stroke-only-preset';

function shape(overrides: Partial<PptxElement> = {}): PptxElement {
	return {
		type: 'shape',
		id: 'shape-1',
		x: 0,
		y: 0,
		width: 200,
		height: 120,
		shapeType: 'rect',
		...overrides,
	} as PptxElement;
}

describe('getStrokeOnlyPresetPaths', () => {
	// The open family: nothing here has a region to fill, so a CSS border on the
	// element box outlines a rectangle the shape does not have.
	// `arc` is not here any more: the ECMA-376 definition is a split layout (an unstroked
	// wedge FILL path plus a fill="none" stroke path) and PowerPoint fills the wedge when
	// the shape has a fill, so it is a split-stroke preset (see below), not an open one.
	const OPEN = ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3'];
	// The closed family: a border on the box (or on its clip-path) is correct.
	const CLOSED = ['rect', 'ellipse', 'roundRect', 'triangle', 'pie', 'chord', 'donut'];

	it.each(OPEN)('reports %s as stroke-only', (shapeType) => {
		expect(getStrokeOnlyPresetPaths(shape({ shapeType } as Partial<PptxElement>))).toBeDefined();
		expect(isStrokeOnlyPresetElement(shape({ shapeType } as Partial<PptxElement>))).toBeTruthy();
	});

	it.each(CLOSED)('leaves %s alone', (shapeType) => {
		expect(getStrokeOnlyPresetPaths(shape({ shapeType } as Partial<PptxElement>))).toBeUndefined();
		expect(isStrokeOnlyPresetElement(shape({ shapeType } as Partial<PptxElement>))).toBeFalsy();
	});

	it('ignores custom geometry, which paints through its own pathData', () => {
		expect(
			getStrokeOnlyPresetPaths(
				shape({ shapeType: 'line', pathData: 'M 0 0 L 5 5' } as Partial<PptxElement>),
			),
		).toBeUndefined();
		expect(
			getStrokeOnlyPresetPaths(shape({ shapeType: 'custom' } as Partial<PptxElement>)),
		).toBeUndefined();
		expect(getStrokeOnlyPresetPaths(shape({ shapeType: undefined }))).toBeUndefined();
	});

	it('never reinterprets a picture as an outline', () => {
		// A picture paints its own bitmap; `line` here would be a crop geometry.
		expect(
			getStrokeOnlyPresetPaths(shape({ type: 'image', shapeType: 'line' } as Partial<PptxElement>)),
		).toBeUndefined();
	});

	it('keeps a horizontal or vertical line straight at zero height or width', () => {
		expect(
			getStrokeOnlyPresetPaths(
				shape({ shapeType: 'line', width: 400, height: 0 } as Partial<PptxElement>),
			),
		).toStrictEqual([{ d: 'M 0 0 L 400 0', fill: 'none', stroke: true }]);
		expect(
			getStrokeOnlyPresetPaths(
				shape({ shapeType: 'line', width: 0, height: 120 } as Partial<PptxElement>),
			),
		).toStrictEqual([{ d: 'M 0 0 L 0 120', fill: 'none', stroke: true }]);
	});
});

describe('strokeOnlyPresetPathData', () => {
	it('joins the stroked sub-paths', () => {
		expect(
			strokeOnlyPresetPathData(
				shape({ shapeType: 'bentConnector3', width: 200, height: 120 } as Partial<PptxElement>),
			),
		).toBe('M 0 0 L 100 0 L 100 120 L 200 120');
	});

	it('is undefined for a closed preset', () => {
		expect(strokeOnlyPresetPathData(shape())).toBeUndefined();
	});
});

describe('splitStrokePresetPathData', () => {
	// PowerPoint (COM export) strokes these only along their stroked sub-paths: an arc is
	// the curve alone (no radial edges), a smiley has no line across the mouth.
	const SPLIT = ['arc', 'smileyFace', 'can', 'cube', 'bevel', 'leftBracket', 'rightBrace'];

	it.each(SPLIT)('reports %s as a split-layout preset', (shapeType) => {
		expect(getSplitStrokePresetPaths(shape({ shapeType }))).toBeDefined();
		expect(getStrokeOnlyPresetPaths(shape({ shapeType }))).toBeUndefined();
		expect(splitStrokePresetPathData(shape({ shapeType }))).toBeTruthy();
	});

	it('strokes an arc as its curve alone, without the wedge radii', () => {
		const d = splitStrokePresetPathData(
			shape({ shapeType: 'arc', shapeAdjustments: { adj1: 16200000, adj2: 0 } }),
		)!;
		expect(d).toContain('A ');
		expect(d).not.toContain('L ');
		expect(d).not.toContain('Z');
	});

	it('leaves ordinary and open presets alone', () => {
		expect(splitStrokePresetPathData(shape({ shapeType: 'rect' }))).toBeUndefined();
		expect(splitStrokePresetPathData(shape({ shapeType: 'line' }))).toBeUndefined();
	});
});
