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

	it('preserves the second rich run and authored formatting for an appended title', () => {
		const source = columnChart();
		const original = source.title!.text!;
		const next = setChartTitle(source, `${original} report`, { preserveRuns: true });
		const oldRuns = source.title!.tx!.rich!.paragraphs[0]!.runs;
		const newRuns = next.title!.tx!.rich!.paragraphs[0]!.runs;
		expect(newRuns).toHaveLength(oldRuns.length);
		expect(newRuns[0]).toEqual(oldRuns[0]);
		expect(newRuns.at(-1)).toEqual({ ...oldRuns.at(-1), text: `${oldRuns.at(-1)!.text} report` });
		expect(next.title!.text).toBe(`${original} report`);
		expect(written(next)).toContain(' 2026 report');
	});

	it('collapses an unrelated title rewrite while preserving the first rich run formatting', () => {
		const next = setChartTitle(columnChart(), 'A completely new title', { preserveRuns: true });
		expect(next.title!.tx!.rich!.paragraphs[0]!.runs).toHaveLength(1);
		expect(written(next)).toContain('<a:rPr lang="en-US" sz="1400" b="1"/>');
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
	it('keeps display-unit labels and source linking across unit edits', () => {
		const source = columnChart();
		source.plotArea.axes[1]!.displayUnits = {
			builtInUnit: 'thousands',
			label: { tx: { value: 'Hundreds' } },
		};
		const next = setChartAxis(source, 'val', {
			displayUnits: 100,
			numberFormat: '0.00',
			numberFormatSourceLinked: true,
		});
		expect(next.plotArea.axes[1]!.displayUnits).toEqual({
			customUnit: 100,
			label: { tx: { value: 'Hundreds' } },
		});
		expect(next.plotArea.axes[1]!.numberFormat?.sourceLinked).toBe(true);
		expect(source.plotArea.axes[1]!.displayUnits?.builtInUnit).toBe('thousands');
		const parsed = parseChartSpace(written(next)).chartSpace;
		expect(parsed?.plotArea.axes[1]!.displayUnits?.customUnit).toBe(100);
		expect(() => setChartAxis(source, 'val', { displayUnits: Infinity })).toThrow(RangeError);
	});

	it('writes category label spacing and validates its OOXML bounds', () => {
		const next = setChartAxis(columnChart(), 'cat', {
			labelAlign: 'l',
			labelOffset: 0,
			tickLabelSkip: 2,
			tickMarkSkip: 3,
			auto: false,
			noMultiLevelLabels: true,
		});
		const xml = written(next);
		expect(xml).toContain('<c:lblOffset val="0"/>');
		expect(xml).toContain('<c:tickLblSkip val="2"/>');
		expect(() => setChartAxis(next, 'cat', { tickLabelSkip: 0 })).toThrow(RangeError);
		expect(() => setChartAxis(next, 'cat', { labelOffset: 1001 })).toThrow(RangeError);
		expect(
			setChartAxis(next, 'cat', { tickLabelSkip: null }).plotArea.axes[0]!.tickLabelSkip,
		).toBeUndefined();
	});
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
