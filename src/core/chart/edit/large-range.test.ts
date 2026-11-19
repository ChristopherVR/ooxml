import { describe, expect, it } from 'vitest';
import { chartExSpaceWorkbookWrites } from '../ex/workbook-writes';
import { parseChartExSpace } from '../ex/parse';
import { writeChartExSpaceOverSource } from '../ex/keep-source';
import { writeChartSpaceOverSource } from '../keep-source';
import { parseChartSpace } from '../parse-space';
import {
	CHART_POINTS_LIMITED,
	chartPointBudget,
	slottedPoints,
	type ChartPointCache,
} from '../point-limit';
import { chartSpaceWorkbookWrites } from '../workbook/space-writes';
import { applyChartExSpaceDataSheetEdits, chartExSpaceDataSheet } from './chartex-sheet';
import { applyChartSpaceDataSheetEdits, chartSpaceDataSheet } from './data-sheet';
import { setChartSeriesValues } from './series';

// Whole-column charts: every series declares a range of a million cells and holds a few points,
// which is how Excel saves a chart over `A:D`.
const FULL = 1_048_575;
const C =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const CX =
	'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const COLUMNS = 'BCDEF';

const points = (slots: readonly number[]) =>
	slots.map((slot) => `<c:pt idx="${slot}"><c:v>${slot + 1}</c:v></c:pt>`).join('');
const series = (index: number, slots: readonly number[]) => {
	const column = COLUMNS[index]!;
	return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:val><c:numRef><c:f>S!$${column}$2:$${column}$1048576</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${FULL}"/>${points(slots)}</c:numCache></c:numRef></c:val></c:ser>`;
};
const chart = (body: string) =>
	`<c:chartSpace ${C}><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${body}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

/** Five series of three points; the last has a fourth, the third a point in the last row. */
const FAR = FULL - 1;
const slotsOf = (index: number) =>
	index === 4 ? [0, 1, 2, 3] : index === 2 ? [0, 1, 2, FAR] : [0, 1, 2];
const xml = chart([0, 1, 2, 3, 4].map((i) => series(i, slotsOf(i))).join(''));
const load = () => parseChartSpace(xml);

const values = (space: ReturnType<typeof load>['chartSpace'], index: number) =>
	space.plotArea.groups[0]!.series[index]!.values!.cache!;
const slots = (cache: ChartPointCache) => slottedPoints(cache).map(([slot]) => slot);
const valueAt = (cache: ChartPointCache<{ index?: number; value: string }>, slot: number) =>
	slottedPoints(cache).find(([at]) => at === slot)?.[1].value;

describe('editing a chart with a large declared range', () => {
	it('cuts the views of this chart, so the cases below are the real ones', () => {
		const { chartSpace, issues } = load();
		expect(issues.map((issue) => issue.code)).toContain(CHART_POINTS_LIMITED);
		expect(values(chartSpace, 0).viewLength).toBe(3);
		expect(values(chartSpace, 4).viewLength).toBe(4);
		expect(chartSpaceDataSheet(chartSpace).categories).toHaveLength(4);
	});

	it('edits the first of five series when one is a point longer', () => {
		const five = chart(
			[0, 1, 2, 3, 4].map((i) => series(i, i === 4 ? [0, 1, 2, 3] : [0, 1, 2])).join(''),
		);
		const { chartSpace } = parseChartSpace(five);
		expect(chartSpaceDataSheet(chartSpace).categories).toHaveLength(4);
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'value', series: 0, category: 0, value: 9 },
		]);
		expect(valueAt(values(edited, 0), 0)).toBe('9');
		for (const index of [0, 1, 2, 3, 4]) expect(values(edited, index).pointCount).toBe(FULL);
	});

	it('edits a value of a series shorter than the sheet', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'value', series: 0, category: 1, value: 42 },
		]);
		const cache = values(edited, 0);
		expect(cache.pointCount).toBe(FULL);
		expect(valueAt(cache, 1)).toBe('42');
		expect(slots(cache)).toEqual([0, 1, 2]);
		expect(chartSpaceDataSheet(edited).series[0]!.values.slice(0, 4)).toEqual([1, 42, 3, null]);
		expect(writeChartSpaceOverSource(xml, edited)).toContain(`<c:ptCount val="${FULL}"/>`);
	});

	it('edits a row the series does not reach yet, and keeps a point past the view', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'value', series: 2, category: 3, value: 7 },
		]);
		const cache = values(edited, 2);
		expect(slots(cache)).toEqual([0, 1, 2, 3, FAR]);
		expect(valueAt(cache, FAR)).toBe(String(FAR + 1));
		expect(cache.pointCount).toBe(FULL);
	});

	it('adds a category in the row after the data, inside the declared range', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [{ kind: 'add-category' }]);
		for (const index of [0, 1, 2, 3, 4]) {
			const cache = values(edited, index);
			expect(cache.pointCount).toBe(FULL);
			expect(valueAt(cache, 4)).toBe('0');
		}
		expect(valueAt(values(edited, 2), FAR)).toBe(String(FAR + 1));
		expect(chartSpaceDataSheet(edited).categories).toHaveLength(5);
	});

	it('keeps a point already in the row a new category takes', () => {
		const part = chart(series(0, [0, 1]) + series(1, [0, 2]));
		const { chartSpace } = parseChartSpace(part, { pointBudget: chartPointBudget(0) });
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [{ kind: 'add-category' }]);
		expect(valueAt(values(edited, 1), 2)).toBe('3');
		expect(valueAt(values(edited, 0), 2)).toBe('0');
	});

	it('removes a category from every series, moving later points up in the same range', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'remove-category', category: 1 },
		]);
		const cut = values(edited, 2);
		expect(cut.pointCount).toBe(FULL);
		expect(slots(cut)).toEqual([0, 1, FAR - 1]);
		expect(valueAt(cut, 1)).toBe('3');
		expect(valueAt(cut, FAR - 1)).toBe(String(FAR + 1));
		expect(chartSpaceDataSheet(edited).categories).toHaveLength(3);
		const saved = parseChartSpace(writeChartSpaceOverSource(xml, edited)).chartSpace;
		expect(values(saved, 2).pointCount).toBe(FULL);
		expect(slots(values(saved, 2))).toEqual([0, 1, FAR - 1]);
	});

	it('removes a row only the padding of the sheet shows', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'remove-category', category: 3 },
		]);
		expect(slots(values(edited, 0))).toEqual([0, 1, 2]);
		expect(slots(values(edited, 4))).toEqual([0, 1, 2]);
	});

	it('moves points past the view when a whole shorter vector replaces it', () => {
		const { chartSpace } = load();
		const cache = values(setChartSeriesValues(chartSpace, 2, [9, 9, 9]), 2);
		expect(cache.pointCount).toBe(FULL);
		expect(slots(cache)).toEqual([0, 1, 2, FAR - 1]);
	});
});

describe('a ChartEx chart with a large declared range', () => {
	const lvl = (count: number) =>
		`<cx:lvl ptCount="${FULL}" formatCode="General">${Array.from({ length: count }, (_, i) => `<cx:pt idx="${i}">${i + 1}</cx:pt>`).join('')}<cx:pt idx="${FAR}">9</cx:pt></cx:lvl>`;
	const data = (id: number, count: number) =>
		`<cx:data id="${id}"><cx:numDim type="val"><cx:f>S!$${COLUMNS[id]}$2:$${COLUMNS[id]}$1048576</cx:f>${lvl(count)}</cx:numDim></cx:data>`;
	const ids = [0, 1, 2, 3, 4];
	const part = `<cx:chartSpace ${CX}><cx:chartData>${ids.map((i) => data(i, i === 4 ? 4 : 3)).join('')}</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion>${ids.map((i) => `<cx:series layoutId="clusteredColumn"><cx:dataId val="${i}"/></cx:series>`).join('')}</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
	const level = (space: ReturnType<typeof parseChartExSpace>['chartExSpace'], id: number) =>
		space.chartData.data[id]!.dimensions[0]!.levels[0]!;

	it('edits a value of a shorter series and keeps its range', () => {
		const { chartExSpace } = parseChartExSpace(part);
		expect(chartExSpaceDataSheet(chartExSpace).categories).toHaveLength(5);
		const edited = applyChartExSpaceDataSheetEdits(chartExSpace, [
			{ kind: 'value', series: 0, category: 0, value: 5 },
		]);
		expect(level(edited, 0).pointCount).toBe(FULL);
		expect(valueAt(level(edited, 0), 0)).toBe('5');
		expect(valueAt(level(edited, 0), FAR)).toBe('9');
		expect(writeChartExSpaceOverSource(part, edited)).toContain(`ptCount="${FULL}"`);
	});

	it('adds and removes a category inside the declared range', () => {
		const { chartExSpace } = parseChartExSpace(part);
		const added = applyChartExSpaceDataSheetEdits(chartExSpace, [{ kind: 'add-category' }]);
		expect(valueAt(level(added, 0), 5)).toBe('0');
		expect(level(added, 0).pointCount).toBe(FULL);
		const removed = applyChartExSpaceDataSheetEdits(chartExSpace, [
			{ kind: 'remove-category', category: 0 },
		]);
		expect(level(removed, 0).pointCount).toBe(FULL);
		expect(valueAt(level(removed, 0), FAR - 1)).toBe('9');
	});

	it('writes only the changed cell of the declared range', () => {
		const { chartExSpace } = parseChartExSpace(part);
		const edited = applyChartExSpaceDataSheetEdits(chartExSpace, [
			{ kind: 'value', series: 1, category: 2, value: 8 },
		]);
		expect(chartExSpaceWorkbookWrites(edited, undefined, chartExSpace)).toEqual([
			{
				formula: 'S!$C$2:$C$1048576',
				isNumeric: true,
				cells: [{ index: 2, value: '8' }],
				cellCount: FULL,
			},
		]);
	});
});

describe('the workbook writes of a chart edit', () => {
	it('name only the cells the edit changed', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'value', series: 3, category: 2, value: 8 },
		]);
		expect(chartSpaceWorkbookWrites(edited, chartSpace)).toEqual([
			{
				formula: 'S!$E$2:$E$1048576',
				isNumeric: true,
				cells: [{ index: 2, value: '8' }],
				cellCount: FULL,
			},
		]);
	});

	it('move a point past the view with the row removed before it', () => {
		const { chartSpace } = load();
		const edited = applyChartSpaceDataSheetEdits(chartSpace, [
			{ kind: 'remove-category', category: 0 },
		]);
		const write = chartSpaceWorkbookWrites(edited, chartSpace).find(({ formula }) =>
			formula.startsWith('S!$D$'),
		)!;
		expect(write.cellCount).toBe(FULL);
		expect(write.cells).toEqual(
			expect.arrayContaining([
				{ index: FAR - 1, value: String(FAR + 1) },
				{ index: FAR, value: '' },
				{ index: 2, value: '' },
			]),
		);
	});
});
