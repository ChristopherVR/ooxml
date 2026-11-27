import {
	MAX_COL,
	MAX_ROW,
	rangesIntersect,
	rangeContains,
	parseAddress,
	type CellAddress,
	type CellRange,
} from '../address';
import { forEachCellInRange, putCell, deleteCell } from '../cells';
import type { PivotTable, PivotTableOptions, Workbook } from '../model';
import { pivotGrid } from '../pivot/aggregate';
import { pivotCaption, pivotSourceData, validatePivotFields } from '../pivot/data';
import { pivotSourceRange } from '../pivot/source';
import { type EditContext, sheetAt } from './context';
import { stylePatcher } from './format';
import { styleAt } from '../styles';
import { authoredCaptions } from '../pivot/captions';
import { normalizedPivotMeasures } from '../pivot/measures';
import { pivotDateFieldsOf, roundDateRecords } from '../pivot/dates';
import { reconcileDateAxes } from '../pivot/date-axes';
import { normalizedFieldSubtotals } from '../pivot/subtotals';
import { applyPivotChartUpdates, detachPivotCharts, pivotChartUpdates } from './pivot-charts';
import {
	pivotHasSlicers,
	removeSlicers,
	removeTimelines,
	slicerOnPivot,
	timelineOnPivot,
} from '../slicer/book';

export function pivotAt(ctx: EditContext, sheet: number, name: string): PivotTable {
	const pivot = sheetAt(ctx.workbook, sheet).pivotTables?.find((entry) => entry.name === name);
	if (!pivot) throw new Error('The selected PivotTable is not available.');
	return pivot;
}

function uniqueName(ctx: EditContext): string {
	const used = reservedPivotNames(ctx.workbook);
	for (let i = 1; ; i++) if (!used.has(`pivottable${i}`)) return `PivotTable${i}`;
}

export function reservedPivotNames(workbook: Workbook): Set<string> {
	return new Set(
		workbook.sheets
			.flatMap((sheet) => [
				...(sheet.pivotTables ?? []).map((pivot) => pivot.name),
				...(sheet.preserved.get('source:unsupportedPivotNames') ?? []),
			])
			.map((name) => name.toLowerCase()),
	);
}

export function checkName(ctx: EditContext, name: string, previous?: string): void {
	if (
		!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ||
		name.length > 255 ||
		parseAddress(name) ||
		(name.toLowerCase() !== previous?.toLowerCase() &&
			reservedPivotNames(ctx.workbook).has(name.toLowerCase()))
	)
		throw new Error('A PivotTable name must be unique and cannot look like a cell reference.');
}

function overlapsSource(
	ctx: EditContext,
	s: number,
	pivot: PivotTable,
	location: CellRange,
	fromCache: boolean,
) {
	let source: ReturnType<typeof pivotSourceRange>;
	try {
		source = pivotSourceRange(ctx.workbook, s, pivot.source);
	} catch (error) {
		if (!fromCache) throw error;
		return false;
	}
	return source.sheet === s && rangesIntersect(source.range, location);
}

function plan(
	ctx: EditContext,
	s: number,
	pivot: PivotTable,
	old: PivotTable | undefined,
	fromCache: boolean,
) {
	const sheet = sheetAt(ctx.workbook, s);
	const grid = pivotGrid(pivot);
	const start = pivot.location.start;
	const location = {
		start,
		end: {
			row: start.row + grid.values.length - 1,
			col: start.col + (grid.values[0]?.length ?? 0) - 1,
		},
	};
	if (
		![start.row, start.col].every(Number.isInteger) ||
		start.row < 0 ||
		start.col < 0 ||
		location.end.row > MAX_ROW ||
		location.end.col > MAX_COL
	)
		throw new Error('The PivotTable result does not fit on the worksheet.');
	if (
		overlapsSource(ctx, s, pivot, location, fromCache) ||
		sheet.merges.some((range) => rangesIntersect(range, location)) ||
		sheet.tables.some((table) => rangesIntersect(table.range, location)) ||
		sheet.pivotTables?.some((other) => other !== old && rangesIntersect(other.location, location))
	)
		throw new Error('The PivotTable would overlap source data, another table or merged cells.');
	let occupied = false;
	forEachCellInRange(sheet, location, (cell, row, col) => {
		if (
			(!old || !rangeContains(old.location, { row, col })) &&
			(cell.value !== null || cell.formula !== undefined)
		)
			occupied = true;
	});
	if (occupied) throw new Error('The PivotTable would overwrite existing cell content.');
	return { grid, location };
}

export function writePivot(
	ctx: EditContext,
	s: number,
	pivot: PivotTable,
	old?: PivotTable,
	label?: string,
	inside?: () => void,
	fromCache = false,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const { grid, location } = plan(ctx, s, pivot, old, fromCache);
	const next = { ...pivot, location, dirty: true };
	const charts = old
		? pivotChartUpdates(ctx.workbook, sheet.name, next, grid, old.name)
		: undefined;
	const ranges = [...(old ? [old.location] : []), location];
	ctx.run(
		label ?? (old ? 'Refresh PivotTable' : 'Create PivotTable'),
		'annotations',
		[
			{ kind: 'parts', sheet: s, parts: ['pivotTables'] },
			{ kind: 'cells', sheet: s, ranges },
			...[...(charts?.keys() ?? [])].map((sheet) => ({
				kind: 'parts' as const,
				sheet,
				parts: ['drawings' as const],
			})),
			...(inside ? [{ kind: 'meta' as const }] : []),
		],
		() => {
			const compact = pivot.layout === 'compact';
			const indents = new Map(
				[...new Set([0, ...(grid.indents ?? [])])].map(
					(indent) =>
						[
							indent,
							stylePatcher(ctx.workbook, {
								alignment: { indent, ...(compact || indent > 0 ? { horizontal: 'left' } : {}) },
							}),
						] as const,
				),
			);
			const styles = new Map<string, number>();
			forEachCellInRange(sheet, location, (cell, row, col) => {
				if (cell.styleId !== undefined) styles.set(`${row}:${col}`, cell.styleId);
			});
			if (old)
				forEachCellInRange(sheet, old.location, (cell, row, col) => {
					if (cell.styleId !== undefined)
						putCell(sheet, row, col, { value: null, styleId: cell.styleId });
					else deleteCell(sheet, row, col);
				});
			inside?.();
			const list = sheet.pivotTables ?? [];
			sheet.pivotTables = old
				? list.map((entry) => (entry === old ? next : entry))
				: [...list, next];
			applyPivotChartUpdates(ctx.workbook, charts);
			grid.values.forEach((line, row) =>
				line.forEach((value, col) => {
					const r = startRow(location, row);
					const c = location.start.col + col;
					let styleId = styles.get(`${r}:${c}`);
					if (col < (grid.rowLabelColumns ?? 1)) {
						const indent = col === 0 ? (grid.indents?.[row] ?? 0) : 0;
						const alignment = styleAt(ctx.workbook, styleId).alignment;
						if (
							(alignment?.indent ?? 0) !== indent ||
							(compact && alignment?.horizontal !== 'left')
						)
							styleId = indents.get(indent)!(styleId);
					}
					putCell(sheet, r, c, { value, ...(styleId !== undefined ? { styleId } : {}) });
				}),
			);
		},
		{ sheet: s, ranges },
	);
}
const startRow = (range: CellRange, offset: number) => range.start.row + offset;

export function createPivotTable(
	ctx: EditContext,
	s: number,
	destination: CellAddress,
	options: PivotTableOptions,
): string {
	const data = pivotSourceData(ctx.workbook, s, options.source);
	const name = options.name ?? uniqueName(ctx);
	checkName(ctx, name);
	const aggregate = options.aggregate ?? 'sum';
	const captions = authoredCaptions(options);
	const pivot: PivotTable = {
		...data,
		...(captions ? { captions } : {}),
		name,
		rowField: options.rowField,
		layout: options.layout ?? 'tabular',
		subtotalPosition:
			options.subtotalPosition ??
			(options.layout && options.layout !== 'tabular' ? 'top' : 'bottom'),
		...(options.secondRowField ? { secondRowField: options.secondRowField } : {}),
		...(options.secondRowField && options.moreRowFields?.length
			? { moreRowFields: [...options.moreRowFields] }
			: {}),
		...(options.fieldSubtotals?.length
			? { fieldSubtotals: normalizedFieldSubtotals(options.fieldSubtotals) }
			: {}),
		...(options.dateGroups?.length ? { dateGroups: structuredClone(options.dateGroups) } : {}),
		...(ctx.workbook.date1904 ? { date1904: true } : {}),
		...(options.subtotals !== undefined ? { subtotals: options.subtotals } : {}),
		...(options.repeatLabels !== undefined ? { repeatLabels: options.repeatLabels } : {}),
		valueField: options.valueField,
		aggregate,
		valueCaption: pivotCaption(aggregate, options.valueField, { captions }),
		...(typeof options.valueNumberFormat === 'string' && options.valueNumberFormat
			? { valueNumberFormat: options.valueNumberFormat }
			: {}),
		...(options.pageFilter !== undefined
			? { pageFilter: structuredClone(options.pageFilter) }
			: {}),
		...(options.fieldGroups?.length ? { fieldGroups: structuredClone(options.fieldGroups) } : {}),
		...(options.valuesAxis !== undefined ? { valuesAxis: options.valuesAxis } : {}),
		...(options.valuesPosition !== undefined ? { valuesPosition: options.valuesPosition } : {}),
		...(options.additionalValues
			? { additionalValues: normalizedPivotMeasures(options.additionalValues, { captions }) }
			: {}),
		location: { start: { ...destination }, end: { ...destination } },
		...(options.columnField ? { columnField: options.columnField } : {}),
		...(options.columnField && options.moreColumnFields?.length
			? { moreColumnFields: [...options.moreColumnFields] }
			: {}),
	};
	if (pivot.dateGroups) reconcileDateAxes(pivot);
	const dates = pivotDateFieldsOf(ctx.workbook, s, pivot);
	if (dates.length) pivot.dateFields = dates;
	roundDateRecords(pivot, dates, ctx.workbook.date1904);
	validatePivotFields(pivot);
	writePivot(ctx, s, pivot);
	return name;
}

export function deletePivotTable(ctx: EditContext, s: number, name: string): void {
	const sheet = sheetAt(ctx.workbook, s);
	const pivot = pivotAt(ctx, s, name);
	const charts = detachPivotCharts(ctx.workbook, sheet.name, name);
	ctx.run(
		'Delete PivotTable',
		'annotations',
		[
			{ kind: 'parts', sheet: s, parts: ['pivotTables'] },
			{ kind: 'cells', sheet: s, ranges: [pivot.location] },
			...[...charts.keys()].map((sheet) => ({
				kind: 'parts' as const,
				sheet,
				parts: ['drawings' as const],
			})),
			...(pivotHasSlicers(ctx.workbook, sheet.sheetId, pivot.name)
				? [{ kind: 'workbook' as const }]
				: []),
		],
		() => {
			removeSlicers(ctx.workbook, (slicer) =>
				slicerOnPivot(ctx.workbook, slicer.cache, sheet.sheetId, pivot.name),
			);
			removeTimelines(ctx.workbook, (timeline) =>
				timelineOnPivot(ctx.workbook, timeline.cache, sheet.sheetId, pivot.name),
			);
			sheet.pivotTables = sheet.pivotTables!.filter((entry) => entry !== pivot);
			applyPivotChartUpdates(ctx.workbook, charts);
			forEachCellInRange(sheet, pivot.location, (_cell, row, col) => deleteCell(sheet, row, col));
		},
		{ sheet: s, ranges: [pivot.location] },
	);
}
