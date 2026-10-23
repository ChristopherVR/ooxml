import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { parseXml } from '../../xml/index';
import { applyChartSpaceDataSheetEdits, chartSpaceDataSheet } from './data-sheet';
import { DEFAULT_CHART_DATA_SHEET, newChartSpace } from './new-chart';

const reload = (xml: string) => parseChartSpace(parseXml(xml).documentElement!).chartSpace;

describe('the data sheet of a chart part', () => {
	it('reads a new chart back as the sheet it was built from', () => {
		expect(chartSpaceDataSheet(newChartSpace())).toEqual(DEFAULT_CHART_DATA_SHEET);
	});

	it('keeps series formatting and a name reference through edits', () => {
		const base = newChartSpace();
		base.plotArea.groups[0]!.series[1]!.spPr = {
			fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } },
		};
		const edited = applyChartSpaceDataSheetEdits(base, [
			{ kind: 'series-name', series: 1, name: 'Profit' },
			{ kind: 'remove-series', series: 0 },
			{ kind: 'remove-category', category: 3 },
		]);
		const [first] = edited.plotArea.groups[0]!.series;
		expect(first!.spPr?.fill).toMatchObject({ kind: 'solid' });
		expect(first!.tx?.reference?.formula).toBe('Sheet1!$B$1');
		expect(first!.values?.formula).toBe('Sheet1!$B$2:$B$4');
		expect(first!.categories?.formula).toBe('Sheet1!$A$2:$A$4');
		expect(chartSpaceDataSheet(edited)).toEqual({
			categories: ['Category 1', 'Category 2', 'Category 3'],
			series: [
				{ name: 'Profit', values: [2.4, 4.4, 1.8] },
				{ name: 'Series 3', values: [2, 2, 3] },
			],
		});
	});

	it('survives writing and reading the part', () => {
		const edited = applyChartSpaceDataSheetEdits(newChartSpace(), [
			{ kind: 'value', series: 2, category: 0, value: 12 },
			{ kind: 'add-series' },
		]);
		expect(chartSpaceDataSheet(reload(writeChartSpace(edited)))).toEqual(
			chartSpaceDataSheet(edited),
		);
	});

	it('ignores edits the sheet refuses and returns the same model', () => {
		const space = newChartSpace();
		expect(
			applyChartSpaceDataSheetEdits(space, [{ kind: 'value', series: 9, category: 0, value: 1 }]),
		).toBe(space);
	});
});

describe('a new chart part', () => {
	it('switches to the picked family through the type switch', () => {
		const pie = newChartSpace(undefined, { kind: 'pie' }, { title: 'Sales' });
		expect(pie.plotArea.groups[0]!.element).toBe('pieChart');
		expect(pie.plotArea.axes).toEqual([]);
		expect(pie.title?.text).toBe('Sales');
		expect(pie.legend?.position).toBe('b');
		const bars = newChartSpace(undefined, {
			kind: 'bar',
			barDirection: 'bar',
			grouping: 'stacked',
		});
		expect(bars.plotArea.groups[0]).toMatchObject({ barDirection: 'bar', grouping: 'stacked' });
		expect(bars.autoTitleDeleted).toBe(true);
	});

	it('writes a part that reads back unchanged', () => {
		const line = newChartSpace(undefined, { kind: 'line' }, { title: 'Trend' });
		const xml = writeChartSpace(line);
		expect(xml).toContain('<c:lineChart>');
		expect(chartSpaceDataSheet(reload(xml))).toEqual(DEFAULT_CHART_DATA_SHEET);
	});
});
