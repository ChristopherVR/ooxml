import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { chartDataLabelsRange, setChartDataLabelsRange } from './label-range';
import { setChartDataLabels } from './legend-labels';
import { columnChart } from './test-support/column-chart';

describe('value-from-cells labels', () => {
	it('preserves sibling extensions and point overrides through edit and removal', () => {
		const source = columnChart();
		source.plotArea.groups[0]!.dataLabels = {
			labels: [{ index: 1, deleted: true }],
			extLst:
				'<c:extLst><c:ext uri="kept"><z:item xmlns:z="urn:test" value="opaque"/></c:ext><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"><c15:xForSave xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" val="1"/></c:ext></c:extLst>',
		};
		const next = setChartDataLabelsRange(source, {
			formula: 'Sheet1!$D$2:$D$3',
			cache: ['A & B', 'Second'],
		});
		const xml = writeChartSpace(next);
		expect(xml).toContain('A &amp; B');
		expect(xml).toContain('value="opaque"');
		expect(xml).toContain('xForSave');
		const reloaded = parseChartSpace(xml).chartSpace;
		expect(chartDataLabelsRange(reloaded.plotArea.groups[0]!.dataLabels!)).toEqual({
			formula: 'Sheet1!$D$2:$D$3',
			cache: ['A & B', 'Second'],
			show: true,
		});
		expect(reloaded.plotArea.groups[0]!.dataLabels!.labels[0]).toMatchObject({
			index: 1,
			deleted: true,
		});
		const removed = writeChartSpace(setChartDataLabelsRange(next, null));
		expect(removed).not.toContain('datalabelsRange');
		expect(removed).toContain('xForSave');
		expect(removed).toContain('value="opaque"');
		expect(source.plotArea.groups[0]!.dataLabels!.extLst).not.toContain('dlblRangeCache');
	});

	it('keeps identical source extension bytes and targets one series', () => {
		const range = { formula: 'Sheet1!$D$2:$D$3', cache: ['One', 'Two'] };
		const first = setChartDataLabelsRange(columnChart(), range, false, { series: 0 });
		expect(first.plotArea.groups[0]!.dataLabels).toBeUndefined();
		const labels = first.plotArea.groups[0]!.series[0]!.dataLabels!;
		expect(chartDataLabelsRange(labels)?.show).toBe(false);
		const second = setChartDataLabelsRange(first, range, false, { series: 0 });
		expect(second.plotArea.groups[0]!.series[0]!.dataLabels!.extLst).toBe(labels.extLst);
	});

	it('can hide and reshow labels without discarding series-specific formatting', () => {
		const source = setChartDataLabels(columnChart(), { show: true }, { series: 0 });
		const hidden = setChartDataLabels(source, { show: false, preserveHiddenContent: true });
		expect(hidden.plotArea.groups[0]!.dataLabels).toEqual({ labels: [], deleted: true });
		expect(hidden.plotArea.groups[0]!.series[0]!.dataLabels).toEqual(
			source.plotArea.groups[0]!.series[0]!.dataLabels,
		);
		const shown = setChartDataLabels(hidden, { show: true });
		expect(shown.plotArea.groups[0]!.dataLabels!.deleted).toBeUndefined();
		expect(shown.plotArea.groups[0]!.series[0]!.dataLabels).toEqual(
			source.plotArea.groups[0]!.series[0]!.dataLabels,
		);
	});
});
