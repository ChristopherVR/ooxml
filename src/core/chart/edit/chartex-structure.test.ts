import { beforeAll, describe, expect, it } from 'vitest';
import { chartExParts, CHART_EX_FIXTURES } from '../ex/test-support/fixtures';
import { parseChartExSpace } from '../ex/parse';
import { writeChartExSpace } from '../ex/write';
import { writeChartExSpaceOverSource } from '../ex/keep-source';
import { setChartExTitle } from './chartex';
import type { ChartExSpace } from '../ex/model';
import { chartExSpaceWorkbookWrites } from '../ex/workbook-writes';
import {
	addChartExSeries,
	removeChartExSeries,
	setChartExAxis,
	setChartExLayout,
	setChartExPointLabel,
} from './chartex-structure';
import { applyChartExSpaceDataSheetEdits, chartExSpaceDataSheet } from './chartex-sheet';

const models = new Map<string, ChartExSpace>();
beforeAll(async () => {
	for (const [, xml] of await chartExParts(CHART_EX_FIXTURES[1]![0])) {
		const space = parseChartExSpace(xml).chartExSpace;
		models.set(space.plotArea.series[0]!.layoutId, space);
	}
});
const of = (kind: string) => structuredClone(models.get(kind)!);
const reload = (space: ChartExSpace) => parseChartExSpace(writeChartExSpace(space)).chartExSpace;

describe('ChartEx structural and data sheet edits', () => {
	it('retains parent hierarchy levels through leaf and value edits', () => {
		const space = of('treemap');
		const before = chartExSpaceDataSheet(space);
		const edited = applyChartExSpaceDataSheetEdits(space, [
			{ kind: 'category-label', category: 1, label: 'New leaf' },
			{ kind: 'value', series: 0, category: 1, value: 77 },
		]);
		const after = chartExSpaceDataSheet(reload(edited));
		expect(after.categoryLevels.slice(1)).toEqual(before.categoryLevels.slice(1));
		expect(after.categories[1]).toBe('New leaf');
		expect(after.series[0]!.values[1]).toBe(77);
		expect(chartExSpaceDataSheet(space)).toEqual(before);
	});
	it('moves every hierarchy level and point index when deleting a row', () => {
		const space = of('treemap');
		space.plotArea.series[0]!.dataPoints = [{ index: 1 }, { index: 3 }];
		space.plotArea.series[0]!.dataLabels = {
			labels: [{ index: 3, visibility: { value: true } }],
			hidden: [1, 4],
		};
		const before = chartExSpaceDataSheet(space);
		const edited = reload(
			applyChartExSpaceDataSheetEdits(space, [{ kind: 'remove-category', category: 1 }]),
		);
		const after = chartExSpaceDataSheet(edited);
		expect(after.categoryLevels).toEqual(
			before.categoryLevels.map((level) => level.filter((_, i) => i !== 1)),
		);
		expect(edited.plotArea.series[0]!.dataPoints).toEqual([{ index: 2 }]);
		expect(edited.plotArea.series[0]!.dataLabels?.labels[0]?.index).toBe(2);
		expect(edited.plotArea.series[0]!.dataLabels?.hidden).toEqual([3]);
	});
	it('adds independent series data and removes only unreferenced data', () => {
		const space = of('clusteredColumn');
		const added = addChartExSeries(space, 'Other');
		expect(added.plotArea.series).toHaveLength(2);
		expect(added.chartData.data).toHaveLength(space.chartData.data.length + 1);
		expect(added.plotArea.series[1]!.dataId).not.toBe(space.plotArea.series[0]!.dataId);
		expect(added.chartData.data.at(-1)!.dimensions.every((d) => d.formula === undefined)).toBe(
			true,
		);
		const edited = applyChartExSpaceDataSheetEdits(added, [
			{ kind: 'value', series: 1, category: 0, value: 123 },
		]);
		expect(chartExSpaceDataSheet(edited).series[0]!.values[0]).not.toBe(123);
		expect(chartExSpaceDataSheet(edited).series[1]!.values[0]).toBe(123);
		expect(removeChartExSeries(added, 1)).toEqual(space);
		expect(() => removeChartExSeries(space, 0)).toThrow();
		expect(() => addChartExSeries(of('treemap'), 'Other')).toThrow();
	});
	it('switches layout while preserving hierarchy and numeric references', () => {
		const space = of('treemap');
		const before = chartExSpaceDataSheet(space);
		const waterfall = reload(setChartExLayout(space, 'waterfall'));
		expect(waterfall.plotArea.axes).toHaveLength(2);
		expect(waterfall.plotArea.series[0]!.axisIds).toEqual([0, 1]);
		expect(chartExSpaceDataSheet(waterfall)).toEqual(before);
		expect(waterfall.chartData.data[0]!.dimensions.find((d) => d.kind === 'numDim')?.type).toBe(
			'val',
		);
		expect(() => setChartExLayout(space, 'regionMap')).toThrow('geography');
	});
	it('patches axes and individual labels without losing authored sibling metadata', () => {
		const space = of('waterfall');
		const axis = space.plotArea.axes[1]!;
		const edited = setChartExPointLabel(
			setChartExAxis(space, 1, {
				hidden: true,
				numberFormat: { formatCode: '0.00', sourceLinked: false },
			}),
			0,
			2,
			{ visibility: { value: true }, separator: ' / ' },
		);
		const back = reload(edited);
		expect(back.plotArea.axes[1]).toEqual({
			...axis,
			hidden: true,
			numberFormat: { formatCode: '0.00', sourceLinked: false },
		});
		expect(back.plotArea.series[0]!.dataLabels?.labels).toEqual([
			{ index: 2, visibility: { value: true }, separator: ' / ' },
		]);
		expect(() => setChartExAxis(space, 1, { scaling: { kind: 'val', min: NaN } })).toThrow();
	});
});
describe('ChartEx workbook mapping', () => {
	it('maps leaf-first cached hierarchy to outermost-first worksheet columns', () => {
		const space = of('treemap');
		const writes = chartExSpaceWorkbookWrites(space);
		const sheet = chartExSpaceDataSheet(space);
		expect(writes.find((w) => w.formula === 'Sheet1!$T$2:$T$6')?.values).toEqual(
			sheet.categoryLevels[1],
		);
		expect(writes.find((w) => w.formula === 'Sheet1!$U$2:$U$6')?.values).toEqual(
			sheet.categoryLevels[0],
		);
		expect(writes.find((w) => w.formula === 'Sheet1!$V$1')?.values).toEqual([
			sheet.series[0]!.name,
		]);
	});
	it('resolves host defined names but rejects external and mismatched sources', () => {
		const space = of('waterfall');
		const dimension = space.chartData.data[0]!.dimensions[0]!;
		const original = dimension.formula!.formula;
		dimension.formula!.formula = '_xlchart.v1.5';
		expect(() => chartExSpaceWorkbookWrites(space)).toThrow('local');
		expect(
			chartExSpaceWorkbookWrites(space, (name) => (name === '_xlchart.v1.5' ? original : undefined))
				.length,
		).toBeGreaterThan(0);
		expect(() => chartExSpaceWorkbookWrites(of('sunburst'))).toThrow('local');
		dimension.formula!.formula = 'Sheet1!$A$2:$A$99';
		expect(() => chartExSpaceWorkbookWrites(space)).toThrow('cached points');
	});
});

it('preserves source bytes and refuses edits with unknown nested content or invalid values', () => {
	const source = writeChartExSpace(of('waterfall'));
	for (const changed of [
		source.replace('</cx:series>', '<cx:unknown/></cx:series>'),
		source.replace('id="0"', 'id="bad"'),
	]) {
		const parsed = parseChartExSpace(changed);
		expect(parsed.issues.length).toBeGreaterThan(0);
		expect(writeChartExSpaceOverSource(changed, parsed.chartExSpace)).toBe(changed);
		expect(() =>
			writeChartExSpaceOverSource(changed, setChartExTitle(parsed.chartExSpace, 'Unsafe')),
		).toThrow('cannot be preserved');
	}
});
