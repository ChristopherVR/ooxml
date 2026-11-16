import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import {
	chartFilteredSeriesIndexes,
	filterChartSeries,
	removeChartFilteredSeries,
	restoreChartFilteredSeries,
} from './filtered-series';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const ser = (i: number) =>
	`<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${'BCD'[i]}$1</c:f>` +
	`<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>S${i}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
	`<c:val><c:numRef><c:f>Sheet1!$${'BCD'[i]}$2:$${'BCD'[i]}$3</c:f><c:numCache><c:formatCode>General</c:formatCode>` +
	`<c:ptCount val="2"/><c:pt idx="0"><c:v>${i}</c:v></c:pt><c:pt idx="1"><c:v>${i + 1}</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`;
const load = () =>
	parseChartSpace(
		`<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>` +
			`${ser(0)}${ser(1)}${ser(2)}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
			'<c:catAx><c:axId val="1"/><c:scaling/><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
			'<c:valAx><c:axId val="2"/><c:scaling/><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>' +
			'</c:plotArea></c:chart></c:chartSpace>',
	).chartSpace;

describe('chart filtered series', () => {
	it('hides a series into its group extension with its cell references, as Excel writes it', () => {
		const hidden = filterChartSeries(load(), 1);
		expect(hidden.plotArea.groups[0]!.series.map((series) => series.index)).toEqual([0, 2]);
		expect(chartFilteredSeriesIndexes(hidden)).toEqual([1]);
		const xml = writeChartSpace(hidden);
		expect(xml).toMatch(
			/<c:ext uri="\{02D57815-91ED-43cb-92C2-25804820EDAC\}" xmlns:c15="[^"]+"><c15:filteredBarSeries><c15:ser><c:idx val="1"\/>/u,
		);
		expect(xml).toContain(
			'<c:numRef><c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"><c15:formulaRef><c15:sqref>Sheet1!$C$2:$C$3</c15:sqref>',
		);
		expect(parseChartSpace(xml).issues).toEqual([]);
	});

	it('restores a hidden series at its idx position with its formulas', () => {
		const space = parseChartSpace(writeChartSpace(filterChartSeries(load(), 1))).chartSpace;
		const restored = restoreChartFilteredSeries(space, 1);
		const series = restored.plotArea.groups[0]!.series;
		expect(series.map((item) => item.index)).toEqual([0, 1, 2]);
		expect(series[1]!.values?.formula).toBe('Sheet1!$C$2:$C$3');
		expect(restored.plotArea.groups[0]!.extLst).toBeUndefined();
		expect(writeChartSpace(restored)).toBe(writeChartSpace(load()));
	});

	it('retains a combo group with several filtered series and restores each to that group', () => {
		const space = load();
		const bar = space.plotArea.groups[0]!;
		const line = {
			kind: 'line' as const,
			element: 'lineChart',
			is3D: false,
			grouping: 'standard',
			axisIds: [...bar.axisIds],
			series: bar.series.splice(1),
		};
		space.plotArea.groups.push(line);

		const hidden = filterChartSeries(filterChartSeries(space, 2), 1);
		const xml = writeChartSpace(hidden);
		expect(xml.match(/<c15:filteredLineSeries><c15:ser>/gu)).toHaveLength(2);
		const reopened = parseChartSpace(xml).chartSpace;
		expect(reopened.plotArea.groups[1]).toMatchObject({ kind: 'line', series: [] });
		expect(chartFilteredSeriesIndexes(reopened)).toEqual([2, 1]);

		const restored = restoreChartFilteredSeries(restoreChartFilteredSeries(reopened, 2), 1);
		expect(writeChartSpace(restored)).toBe(writeChartSpace(space));
	});

	it('drops a hidden series and refuses what it cannot do', () => {
		const space = filterChartSeries(load(), 2);
		expect(chartFilteredSeriesIndexes(removeChartFilteredSeries(space, 2))).toEqual([]);
		expect(() => removeChartFilteredSeries(space, 0)).toThrow(/No hidden/u);
		expect(() => filterChartSeries(filterChartSeries(space, 1), 0)).toThrow(/last series/u);
	});
});
