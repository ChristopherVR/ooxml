import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml/index';
import { gradientFocusCenter, readDrawingGradient, sortGradientStopEntries } from './gradient-read';
import { DOM_DRAWING_VIEW } from './node-view';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const read = (xml: string) =>
	readDrawingGradient(
		DOM_DRAWING_VIEW,
		parseXml(`<a:gradFill xmlns:a="${A}" ${xml}`).documentElement,
	);

describe('readDrawingGradient', () => {
	it('reads stops, the linear vector, flip and rotation', () => {
		const gradient = read(`flip="XY" rotWithShape="1"><a:gsLst>
			<a:gs pos="100000"><a:srgbClr val="FF0000"/></a:gs>
			<a:gs pos="0"><a:schemeClr val="accent1"/></a:gs>
		</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`);
		expect(gradient.fill).toMatchObject({ kind: 'gradient', angle: 90, scaled: false });
		expect(gradient.fill.stops.map((stop) => stop.position)).toEqual([100, 0]);
		expect(gradient.flip).toBe('xy');
		expect(gradient.rotateWithShape).toBe(true);
		expect(gradient.hasPath).toBe(false);
		expect(sortGradientStopEntries(gradient.stops).map((stop) => stop.position)).toEqual([
			0, 100000,
		]);
	});

	it('keeps stops the neutral fill drops and reads absent rectangle sides as 0', () => {
		const gradient =
			read(`><a:gsLst><a:gs><a:srgbClr val="00FF00"/></a:gs><a:gs pos="5000"/></a:gsLst>
			<a:path path="rect"><a:fillToRect l="50000" t="50000"/></a:path></a:gradFill>`);
		expect(gradient.fill.stops).toEqual([]);
		expect(gradient.stops).toHaveLength(2);
		expect(gradient.stops[0]!.position).toBeUndefined();
		expect(gradient.stops[1]!.color).toBeUndefined();
		expect(gradient.hasPath).toBe(true);
		expect(gradient.fill.path).toBe('rect');
		expect(gradient.fillToRectSides).toEqual({ l: 0.5, t: 0.5, r: 0, b: 0 });
		expect(gradientFocusCenter(gradient.fillToRectSides!)).toEqual({ x: 0.75, y: 0.75 });
		expect(sortGradientStopEntries(gradient.stops, 0).map((stop) => stop.position)).toEqual([
			undefined,
			5000,
		]);
	});

	it('marks a path gradient without @path and drops an unreadable rectangle from the model', () => {
		const gradient = read(`><a:gsLst/><a:path/><a:tileRect l="x"/></a:gradFill>`);
		expect(gradient.hasPath).toBe(true);
		expect(gradient.fill.path).toBeUndefined();
		expect(gradient.fill.tileRect).toBeUndefined();
		expect(Number.isNaN(gradient.tileRectSides!.l)).toBe(true);
	});
});
