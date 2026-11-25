import type { ChartSpace, ChartPlotGroup } from '../../chart/model';
import type { ChartSpaceSeries } from '../../chart/model-series';
import type { DrawingAnchor, ChartObject, PivotTable, Workbook } from '../model';
import { pivotGrid, type PivotGrid } from '../pivot/aggregate';
import {
	chartSpaceWithPivotButtons,
	chartUsesPivot,
	withoutPivotChartButtons,
} from '../pivot/chart-binding';
import { pivotChartSeries } from '../pivot/chart-series';
import { chartObjectFromSpace } from '../read/chart';
import { chartSpaceFromObject } from '../write/chart-space';
import { addChart } from './charts';
import { sheetAt, sheetDrawingAt, type EditContext } from './context';

function pivotFormats(series: readonly ChartSpaceSeries[]): string {
	return `<c:pivotFmts>${series.map(({ index }) => `<c:pivotFmt><c:idx val="${index}"/></c:pivotFmt>`).join('')}</c:pivotFmts>`;
}

function withNewData(group: ChartPlotGroup, series: ChartSpaceSeries[], old: ChartObject) {
	return series.map((entry, index) => {
		const named = old.series.findIndex((s) => s.name === entry.tx?.value);
		const format = group.series[named < 0 ? index : named];
		const { tx, categories, values, xValues, yValues } = entry;
		const next: ChartSpaceSeries = { ...entry, ...format, index, order: index };
		for (const [key, value] of Object.entries({ tx, categories, values, xValues, yValues })) {
			if (value === undefined) Reflect.deleteProperty(next, key);
			else Reflect.set(next, key, value);
		}
		return next;
	});
}

export function addPivotChart(
	ctx: EditContext,
	sheet: number,
	pivotSheet: number,
	name: string,
	anchor: DrawingAnchor,
): number {
	const ws = sheetAt(ctx.workbook, pivotSheet);
	const pivot = ws.pivotTables?.find((entry) => entry.name === name);
	if (!pivot) throw new Error('The selected PivotTable is not available.');
	const formats = new Set<number>();
	for (const host of ctx.workbook.sheets)
		for (const drawing of host.drawings)
			if (
				drawing.kind === 'chart' &&
				drawing.chartSpace?.pivotSource &&
				chartUsesPivot(drawing, ws.name, name)
			)
				formats.add(drawing.chartSpace.pivotSource.formatId);
	let formatId = 0;
	while (formats.has(formatId)) formatId++;
	const frame: ChartObject = {
		kind: 'chart',
		anchor,
		chartType: 'column',
		showLegend: !!pivot.columnField || !!pivot.additionalValues?.length,
		series: pivotChartSeries(pivot, pivotGrid(pivot), ws.name),
	};
	const space = chartSpaceFromObject(frame);
	space.pivotSource = { name: `[Book.xlsx]${ws.name}!${pivot.name}`, formatId };
	space.pivotFormatsXml = pivotFormats(space.plotArea.groups[0]?.series ?? []);
	return addChart(ctx, sheet, { ...frame, chartSpace: chartSpaceWithPivotButtons(space, true) });
}

/** Computes updates before the pivot and its drawings enter the same undo step. */
export function pivotChartUpdates(
	workbook: Workbook,
	sheetName: string,
	pivot: PivotTable,
	grid: PivotGrid,
): Map<number, Map<ChartObject, ChartObject>> {
	const updates = new Map<number, Map<ChartObject, ChartObject>>();
	let data: ChartObject['series'] | undefined;
	workbook.sheets.forEach((sheet, index) => {
		for (const drawing of sheet.drawings) {
			if (
				drawing.kind !== 'chart' ||
				drawing.editBlocked ||
				!chartUsesPivot(drawing, sheetName, pivot.name)
			)
				continue;
			const space = drawing.chartSpace;
			const group = space?.plotArea.groups[0];
			if (!space || !group || space.plotArea.groups.length !== 1) continue;
			data ??= pivotChartSeries(pivot, grid, sheetName);
			const fresh =
				chartSpaceFromObject({ ...drawing, series: data }).plotArea.groups[0]?.series ?? [];
			let next: ChartSpace = {
				...space,
				plotArea: {
					...space.plotArea,
					groups: [{ ...group, series: withNewData(group, fresh, drawing) }],
				},
			};
			if (fresh.length !== group.series.length || !space.pivotFormatsXml)
				next = { ...next, pivotFormatsXml: pivotFormats(fresh) };
			let entries = updates.get(index);
			if (!entries) updates.set(index, (entries = new Map()));
			entries.set(drawing, chartObjectFromSpace(next, drawing));
		}
	});
	return updates;
}

export function showPivotChartButtons(
	ctx: EditContext,
	sheet: number,
	index: number,
	visible: boolean,
): void {
	const drawing = sheetDrawingAt(ctx.workbook, sheet, index);
	if (drawing?.kind !== 'chart' || !drawing.chartSpace?.pivotSource || drawing.editBlocked)
		throw new Error('The selected PivotTable is not available.');
	const next = chartObjectFromSpace(
		chartSpaceWithPivotButtons(drawing.chartSpace, visible),
		drawing,
	);
	ctx.run(
		'Field Buttons',
		'annotations',
		[{ kind: 'parts', sheet, parts: ['drawings'] }],
		() => {
			const ws = sheetAt(ctx.workbook, sheet);
			ws.drawings = ws.drawings.map((entry) => (entry === drawing ? next : entry));
		},
		{ sheet },
	);
}

/** Excel leaves a regular chart with cached data when its report is deleted. */
export function detachPivotCharts(
	workbook: Workbook,
	sheetName: string,
	pivotName: string,
	removedSheet?: number,
) {
	const updates = new Map<number, Map<ChartObject, ChartObject>>();
	workbook.sheets.forEach((sheet, index) => {
		if (index === removedSheet) return;
		for (const chart of sheet.drawings) {
			if (
				chart.kind !== 'chart' ||
				!chart.chartSpace ||
				!chartUsesPivot(chart, sheetName, pivotName)
			)
				continue;
			if (chart.editBlocked) throw new Error(chart.editBlocked);
			const next = structuredClone(chart.chartSpace);
			delete next.pivotSource;
			delete next.pivotFormatsXml;
			withoutPivotChartButtons(next);
			const data = chart.series.map(
				({ nameRef: _name, categoriesRef: _category, valuesRef: _value, ...series }) => series,
			);
			const literal = chartSpaceFromObject({ ...chart, series: data });
			const group = next.plotArea.groups[0];
			if (group && next.plotArea.groups.length === 1) {
				group.series = withNewData(group, literal.plotArea.groups[0]?.series ?? [], chart);
			}
			let entries = updates.get(index);
			if (!entries) updates.set(index, (entries = new Map()));
			entries.set(chart, chartObjectFromSpace(next, chart));
		}
	});
	return updates;
}

export function applyPivotChartUpdates(
	workbook: Workbook,
	updates?: Map<number, Map<ChartObject, ChartObject>>,
): void {
	for (const [index, replacements] of updates ?? []) {
		const host = sheetAt(workbook, index);
		host.drawings = host.drawings.map((drawing) =>
			drawing.kind === 'chart' ? (replacements.get(drawing) ?? drawing) : drawing,
		);
	}
}
