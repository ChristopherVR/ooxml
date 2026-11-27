import { describe, expect, it } from 'vitest';
import {
	chartFillIntent,
	chartOutlineIntent,
	chartShadowIntent,
	chartTextFormatIntents,
} from './shape-format';
import type { DrawingColor } from '../../drawingml/types';

const red: DrawingColor = { kind: 'srgb', value: 'FF0000', transforms: [] };
const area = { kind: 'chartArea' } as const;

describe('chart Format tab intents', () => {
	it('fills, outlines and shadows one element', () => {
		expect(chartFillIntent(area, red)).toEqual({
			op: 'fill',
			target: area,
			fill: { kind: 'solid', color: red },
		});
		expect(chartFillIntent(area, null)).toMatchObject({ fill: { kind: 'none' } });
		expect(chartOutlineIntent(area, null)).toMatchObject({
			op: 'line',
			line: { fill: { kind: 'none' } },
			merge: true,
		});
		expect(chartShadowIntent(area, true)).toMatchObject({ edit: { kind: 'outerShadow' } });
		expect(chartShadowIntent(area, false)).toMatchObject({
			edit: { kind: 'outerShadow', remove: true },
		});
	});

	it('turns a WordArt style into text edits', () => {
		const intents = chartTextFormatIntents(area, { fill: red, outline: null, shadow: true });
		expect(intents).toHaveLength(2);
		expect(intents[0]).toMatchObject({ op: 'text', edit: { fill: { kind: 'solid' }, line: null } });
		expect(intents[1]).toMatchObject({ op: 'text', edit: { effect: { kind: 'outerShadow' } } });
	});
});
