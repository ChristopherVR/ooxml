import { describe, expect, it } from 'vitest';
import type { SparklineGroup, ThemePalette } from '../model';
import { sparklineScale, sparklineView, type SparklineData } from './sparkline-view';

const round = (n: number) => Math.round(n * 1e6) / 1e6;
const THEME: ThemePalette = { colors: [], majorFont: 'Calibri Light', minorFont: 'Calibri' };

function group(overrides: Partial<SparklineGroup> = {}): SparklineGroup {
	return {
		type: 'line',
		lineWeight: 0.75,
		dateAxis: false,
		markers: false,
		high: false,
		low: false,
		first: false,
		last: false,
		negative: false,
		displayEmptyCellsAs: 'gap',
		displayXAxis: false,
		displayHidden: false,
		rightToLeft: false,
		minAxisType: 'individual',
		maxAxisType: 'individual',
		sparklines: [],
		...overrides,
	};
}

const layout = (g: SparklineGroup, values: (number | null)[], all?: SparklineData[], at = 0) => {
	const data = all ?? [{ values }];
	return sparklineView(g, data[at]!, sparklineScale(g, data, at), THEME);
};

describe('sparkline layout', () => {
	it('draws nothing for an all-empty range but keeps the series colour', () => {
		const view = layout(group(), [null, null, null]);
		expect(view).toMatchObject({ lines: [], markers: [], columns: [], color: '#376092' });
		expect(view.axis).toBeUndefined();
		expect(layout(group({ type: 'column' }), []).columns).toEqual([]);
	});

	it('centres a single value and a flat line', () => {
		expect(layout(group(), [5]).lines).toEqual([[{ x: 0.5, y: 0.5 }]]);
		expect(layout(group(), [2, 2]).lines).toEqual([
			[
				{ x: 0, y: 0.5 },
				{ x: 1, y: 0.5 },
			],
		]);
		const flat = layout(group({ type: 'column' }), [3, 3]).columns;
		expect(flat.map((c) => [c.y, c.h])).toEqual([
			[0, 1],
			[0, 1],
		]);
		// A zero column has no height, so none is drawn.
		expect(layout(group({ type: 'column' }), [3, 0, 1]).columns).toHaveLength(2);
	});

	it('scales a line to its extent and splits it at gaps, spans or zeroes empty cells', () => {
		const values = [0, 10, null, 5];
		expect(layout(group(), values).lines).toEqual([
			[
				{ x: 0, y: 1 },
				{ x: 1 / 3, y: 0 },
			],
			[{ x: 1, y: 0.5 }],
		]);
		expect(layout(group({ displayEmptyCellsAs: 'span' }), values).lines).toHaveLength(1);
		expect(layout(group({ displayEmptyCellsAs: 'zero' }), values).lines[0]![2]).toEqual({
			x: 2 / 3,
			y: 1,
		});
	});

	it('colours negative columns and grows columns from the zero axis', () => {
		const view = layout(
			group({
				type: 'column',
				negative: true,
				displayXAxis: true,
				colorNegative: { rgb: 'FF00FF00' },
				colorAxis: { theme: 4 },
			}),
			[4, -4],
		);
		expect(view.axis).toEqual({ y: 0.5, color: '#4472C4' });
		expect(view.columns.map((c) => ({ ...c, x: round(c.x), w: round(c.w) }))).toEqual([
			{ x: 0.05, y: 0, w: 0.4, h: 0.5, color: '#376092' },
			{ x: 0.55, y: 0.5, w: 0.4, h: 0.5, color: '#00FF00' },
		]);
	});

	it('uses manual limits, clamping values outside them', () => {
		const g = group({ minAxisType: 'custom', maxAxisType: 'custom', manualMin: 0, manualMax: 10 });
		expect(layout(g, [5, 20, -5]).lines[0]!.map((p) => p.y)).toEqual([0.5, 0, 1]);
		const columns = layout({ ...g, type: 'column' }, [5, 10]).columns;
		expect(columns.map((c) => [c.y, c.h])).toEqual([
			[0.5, 0.5],
			[0, 1],
		]);
		// A manual minimum without a value falls back to the sparkline's own extent.
		expect(sparklineScale(group({ minAxisType: 'custom' }), [{ values: [2, 4] }], 0)).toEqual({
			min: 2,
			max: 4,
		});
	});

	it('shares one scale across a group when asked', () => {
		const g = group({ minAxisType: 'group', maxAxisType: 'group' });
		const all = [{ values: [0, 5] }, { values: [5, 10] }];
		expect(sparklineScale(g, all, 0)).toEqual({ min: 0, max: 10 });
		expect(layout(g, [], all, 0).lines[0]!.map((p) => p.y)).toEqual([1, 0.5]);
	});

	it('lays out win/loss columns in halves and skips zeroes', () => {
		const view = layout(
			group({ type: 'stacked', negative: true, displayXAxis: true }),
			[3, -1, 0, 7],
		);
		expect(view.columns.map((c) => [c.y, c.h, c.color])).toEqual([
			[0, 0.5, '#376092'],
			[0.5, 0.5, '#D00000'],
			[0, 0.5, '#376092'],
		]);
		expect(view.axis).toEqual({ y: 0.5, color: '#000000' });
	});

	it('marks points in priority order: markers, negative, high, low, first, last', () => {
		const view = layout(
			group({
				markers: true,
				negative: true,
				high: true,
				low: true,
				last: true,
				colorMarkers: { rgb: '111111' },
				colorNegative: { rgb: '222222' },
				colorHigh: { rgb: '333333' },
				colorLow: { rgb: '444444' },
				colorLast: { rgb: '555555' },
			}),
			[1, -2, 9, 3, -2, 4],
		);
		expect(view.markers.map((m) => m.color)).toEqual([
			'#111111',
			'#444444',
			'#333333',
			'#111111',
			'#444444',
			'#555555',
		]);
		expect(layout(group(), [1, 2]).markers).toEqual([]);
	});

	it('spaces points by date and mirrors right-to-left groups', () => {
		const dated = sparklineView(
			group({ dateAxis: true }),
			{ values: [1, 2, 3], dates: [100, 101, 110] },
			{ min: 1, max: 3 },
			THEME,
		);
		expect(dated.lines[0]!.map((p) => p.x)).toEqual([0, 0.1, 1]);
		// Dates that do not match the values fall back to even spacing.
		const uneven = sparklineView(
			group({ dateAxis: true }),
			{ values: [1, 2, 3], dates: [1] },
			{ min: 1, max: 3 },
			THEME,
		);
		expect(uneven.lines[0]!.map((p) => p.x)).toEqual([0, 0.5, 1]);
		const rtl = layout(group({ rightToLeft: true }), [1, 2]);
		expect(rtl.lines[0]!.map((p) => p.x)).toEqual([1, 0]);
	});

	it('hides the axis when zero is outside the scale', () => {
		expect(layout(group({ displayXAxis: true }), [2, 4]).axis).toBeUndefined();
		expect(layout(group({ displayXAxis: true }), [-2, 2]).axis?.y).toBe(0.5);
	});
});
