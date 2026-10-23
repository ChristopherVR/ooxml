import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { setChartAxis, setChartGridlineShape } from './axes';
import { setChartDataLabels, setChartLegend } from './legend-labels';
import { columnChart } from './test-support/column-chart';
import { setChartTitle, setChartTitleVisible } from './title';

const written = (space: ReturnType<typeof columnChart>) => {
	const xml = writeChartSpace(space);
	expect(parseChartSpace(xml).issues).toEqual([]);
	return xml;
};

describe('title edits', () => {
	it('retexts the first run of a rich title and keeps its formatting', () => {
		const next = setChartTitle(columnChart(), 'Revenue');
		expect(next.title?.text).toBe('Revenue');
		const xml = written(next);
		expect(xml).toContain(
			'<a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>Revenue</a:t></a:r></a:p>',
		);
		expect(xml).not.toContain(' 2026');
	});

	it('removes the title and marks the automatic title deleted', () => {
		const next = setChartTitle(columnChart(), undefined);
		expect(next.title).toBeUndefined();
		expect(next.autoTitleDeleted).toBe(true);
		const shown = setChartTitleVisible(next, true);
		expect(shown).toMatchObject({ title: { overlay: false }, autoTitleDeleted: false });
		const fresh = setChartTitle(next, 'New & bold');
		expect(written(fresh)).toContain('<a:t>New &amp; bold</a:t>');
	});
});

describe('legend and data label edits', () => {
	it('hides, shows and places the legend', () => {
		expect(setChartLegend(columnChart(), { visible: false }).legend).toBeUndefined();
		const hidden = setChartLegend(columnChart(), { visible: false });
		expect(setChartLegend(hidden, { position: 't' }).legend).toEqual({
			position: 't',
			overlay: false,
			entries: [],
		});
		expect(setChartLegend(columnChart(), { overlay: true }).legend?.overlay).toBe(true);
	});

	it('shows labels with the value and Office explicit switches', () => {
		const next = setChartDataLabels(columnChart(), { show: true, position: 'outEnd' });
		expect(next.plotArea.groups[0]!.dataLabels).toMatchObject({
			showValue: true,
			showCategoryName: false,
			position: 'outEnd',
		});
		expect(written(next)).toContain('<c:dLblPos val="outEnd"/>');
		const hidden = setChartDataLabels(next, { show: false });
		expect(hidden.plotArea.groups[0]!.dataLabels?.showValue).toBe(false);
	});

	it('edits one series labels without touching the group', () => {
		const next = setChartDataLabels(columnChart(), { showSeriesName: true }, { series: 1 });
		expect(next.plotArea.groups[0]!.dataLabels).toBeUndefined();
		expect(next.plotArea.groups[0]!.series[1]!.dataLabels?.showSeriesName).toBe(true);
	});
});

describe('axis edits', () => {
	it('sets scaling, format, title and visibility', () => {
		const next = setChartAxis(columnChart(), 'val', {
			min: 0,
			max: 50,
			majorUnit: 10,
			logBase: 10,
			reverse: true,
			numberFormat: '0%',
			title: 'Units',
			minorGridlines: true,
		});
		const axis = next.plotArea.axes[1]!;
		expect(axis.scaling).toEqual({ orientation: 'maxMin', min: 0, max: 50, logBase: 10 });
		expect(axis).toMatchObject({
			majorUnit: 10,
			numberFormat: { formatCode: '0%', sourceLinked: false },
			title: { text: 'Units' },
			minorGridlines: true,
		});
		written(next);
		const cleared = setChartAxis(next, { id: 22 }, { min: null, logBase: null, title: null });
		expect(cleared.plotArea.axes[1]!.scaling).toEqual({ orientation: 'maxMin', max: 50 });
		expect(cleared.plotArea.axes[1]!.title).toBeUndefined();
		expect(setChartAxis(columnChart(), 'cat', { visible: false }).plotArea.axes[0]!.deleted).toBe(
			true,
		);
	});

	it('validates the edit and the target', () => {
		expect(() => setChartAxis(columnChart(), 'val', { min: 10, max: 5 })).toThrow(RangeError);
		expect(() => setChartAxis(columnChart(), 'date', { visible: false })).toThrow(RangeError);
		expect(() => setChartAxis(columnChart(), 'cat', { displayUnits: 'thousands' })).toThrow(
			RangeError,
		);
	});

	it('styles gridlines and turns them on', () => {
		const line = { widthEmu: 9525, fill: { kind: 'none' as const } };
		const next = setChartGridlineShape(columnChart(), 'cat', 'major', { line });
		expect(next.plotArea.axes[0]).toMatchObject({
			majorGridlines: true,
			majorGridlinesSpPr: { line },
		});
		expect(written(next)).toContain('<c:majorGridlines><c:spPr><a:ln w="9525"><a:noFill/>');
	});
});
