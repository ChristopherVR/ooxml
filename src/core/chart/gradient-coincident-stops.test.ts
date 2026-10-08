import { expect, it } from 'vitest';
import { NATIVE_COINCIDENT_STOP_RAMP, spreadCoincidentStops } from './gradient-coincident-stops';
import { buildChartGradientDef } from './gradient-definition';

const red = '#FF0000',
	white = '#FFFFFF';

it('opens coincident stops into the measured one-step native ramp after the shared offset', () => {
	expect(
		spreadCoincidentStops([
			{ offset: 0.5, color: red },
			{ offset: 0.5, color: white },
		]),
	).toEqual([
		{ offset: 0.5, color: red },
		{ offset: 0.5 + NATIVE_COINCIDENT_STOP_RAMP, color: white },
	]);
});

it('never crosses the following stop or the end of the gradient', () => {
	expect(
		spreadCoincidentStops([
			{ offset: 0.5, color: red },
			{ offset: 0.5, color: white },
			{ offset: 0.501, color: red },
		]).map((stop) => stop.offset),
	).toEqual([0.5, 0.501, 0.501]);
	expect(
		spreadCoincidentStops([
			{ offset: 1, color: red },
			{ offset: 1, color: white },
		]).map((stop) => stop.offset),
	).toEqual([1, 1]);
});

it('leaves identical duplicates, distinct offsets and opacity-only differences correct', () => {
	const same = [
		{ offset: 0.3, color: red, opacity: 0.5 },
		{ offset: 0.3, color: red, opacity: 0.5 },
	];
	expect(spreadCoincidentStops(same)).toEqual(same);
	expect(
		spreadCoincidentStops([
			{ offset: 0.3, color: red, opacity: 0.5 },
			{ offset: 0.3, color: red },
		])[1]!.offset,
	).toBe(0.3 + NATIVE_COINCIDENT_STOP_RAMP);
	const distinct = [
		{ offset: 0, color: red },
		{ offset: 1, color: white },
	];
	expect(spreadCoincidentStops(distinct)).toEqual(distinct);
});

it('applies only to linear chart paints, whose coincident edge was measured', () => {
	const stops = [
		{ position: 50, color: red, opacity: 1 },
		{ position: 50, color: white, opacity: 1 },
	];
	const linear = buildChartGradientDef('l', { type: 'linear', angle: 135, scaled: true, stops });
	expect(linear.stops.map((stop) => stop.offset)).toEqual([0.5, 0.5 + NATIVE_COINCIDENT_STOP_RAMP]);
	const radial = buildChartGradientDef('r', { type: 'radial', path: 'circle', stops });
	expect(radial.stops.map((stop) => stop.offset)).toEqual([0.5, 0.5]);
});
