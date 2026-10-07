import { expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { parseDrawingColor } from './drawing-color';
import { drawingColorXml } from './write-color';
import type { DiagramColor } from './types';

for (const color of [
	{ kind: 'srgb', value: 'ABCDEF' },
	{ kind: 'scheme', value: 'accent2' },
	{ kind: 'system', value: 'windowText', fallback: '123456' },
	{ kind: 'preset', value: 'red' },
	{ kind: 'scrgb', value: '50000,60000,70000' },
	{ kind: 'hsl', value: '12000000,50000,60000' },
] as DiagramColor[]) {
	it(`round trips ${color.kind} colors and ordered transforms`, () => {
		const complete = {
			...color,
			transforms: [
				{ name: 'lumMod', value: '50000' },
				{ name: 'lumOff', value: '20000' },
				{ name: 'alpha', value: '80000' },
				{ name: 'inv', value: '' },
			],
		};
		expect(parseDrawingColor(parseXml(drawingColorXml(complete)).documentElement)).toEqual(
			complete,
		);
	});
}
