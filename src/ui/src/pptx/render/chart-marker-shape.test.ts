import { expect, it } from 'vitest';
import { buildMarkerPrimitive as coreMarker } from 'ooxml-core/chart';
import { buildMarkerPrimitive } from './chart-marker-shape';

it('delegates to shared marker geometry and retains PowerPoint selection metadata', () => {
	const input = {
		symbol: 'triangle' as const,
		size: 12,
		cx: 10,
		cy: 20,
		fill: '#123456',
		defaultRadius: 3,
		part: { role: 'dataPoint' as const, seriesIndex: 1, pointIndex: 7 },
	};
	expect(buildMarkerPrimitive(input)).toStrictEqual(coreMarker(input));
	expect(buildMarkerPrimitive(input)?.part).toBe(input.part);
});
