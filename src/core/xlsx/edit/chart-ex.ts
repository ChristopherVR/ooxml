// ChartEx uses the existing workbook edit transaction and cell writer, including undo and redo.
import {
	chartExDataChanged,
	chartExSpaceDataSheet,
	chartExSpaceWorkbookWrites,
	chartExLocalRange,
	chartExSummaryFromModel,
	hydrateChartExSpace,
	type ChartExSpace,
} from '../../chart';
import { formatCellAddress } from '../../chart/workbook/cellref';
import { getCell } from '../cells';
import type { UnsupportedDrawingObject, Workbook } from '../model';
import { ensureCell, sheetDrawingAt, type EditContext } from './context';
import { syncTableHeader, writeValue } from './cell-values';

function at(
	book: Workbook,
	s: number,
	index: number,
): UnsupportedDrawingObject & { chartExSpace: ChartExSpace } {
	const drawing = sheetDrawingAt(book, s, index);
	if (drawing?.kind !== 'unsupported' || !drawing.chartExSpace || !drawing.chartExPartName)
		throw new Error('No extended chart at this drawing index.');
	if (drawing.chartExIssues?.length)
		throw new Error('This extended chart contains unsupported content.');
	return drawing as UnsupportedDrawingObject & { chartExSpace: ChartExSpace };
}
function reader(book: Workbook, scope: number) {
	const resolve = (name: string) =>
		book.definedNames.find(
			(n) => n.name.toLowerCase() === name.toLowerCase() && n.localSheet === scope,
		)?.formula ??
		book.definedNames.find(
			(n) => n.name.toLowerCase() === name.toLowerCase() && n.localSheet === undefined,
		)?.formula;
	const target = (formula: string) => {
		const range = chartExLocalRange(formula, resolve);
		const s = book.sheets.findIndex(
			(ws) => ws.name.toLowerCase() === range.sheetName.toLowerCase(),
		);
		if (s < 0) throw new Error('The extended chart source worksheet is missing.');
		return { range, s };
	};
	const read = (formula: string) => {
		const { range, s } = target(formula);
		const rows = range.end.row - range.start.row + 1,
			columns = range.end.col - range.start.col + 1;
		if (rows * columns > 100000)
			throw new Error('The extended chart source range exceeds the editing limit.');
		const values = Array.from({ length: rows * columns }, (_, i) => {
			const cell = getCell(
				book.sheets[s]!,
				range.start.row + Math.floor(i / columns),
				range.start.col + (i % columns),
			);
			const value = cell?.value ?? null;
			if (typeof value === 'object')
				throw new Error('The extended chart source contains an error value.');
			return value;
		});
		return { rows, columns, values };
	};
	return { read, target, resolve };
}
/** Current data, resolved from local sheet cells and hidden chart defined names. */
export function xlsxChartExDataSheet(book: Workbook, s: number, index: number) {
	const chart = at(book, s, index);
	return chartExSpaceDataSheet(hydrateChartExSpace(chart.chartExSpace, reader(book, s).read));
}
/** Apply one neutral operation, synchronizing changed caches and source cells as one undo step. */
export function editChartEx(
	ctx: EditContext,
	s: number,
	index: number,
	operation: (space: ChartExSpace) => ChartExSpace,
): void {
	const chart = at(ctx.workbook, s, index);
	const host = reader(ctx.workbook, s);
	let before = chart.chartExSpace;
	try {
		before = hydrateChartExSpace(before, host.read);
	} catch {
		/* Presentation-only edits can keep unavailable source data untouched. */
	}
	const next = operation(before);
	const changed = chartExDataChanged(before, next);
	const writes = changed ? chartExSpaceWorkbookWrites(next, host.resolve) : [];
	if (changed && next.chartData.externalData)
		throw new Error('A worksheet chart cannot edit an external workbook source.');
	if (
		changed &&
		next.plotArea.series.some((series) =>
			(
				series.data ?? next.chartData.data.find((data) => data.id === series.dataId)
			)?.dimensions.some((d) => !d.formula),
		)
	)
		throw new Error('New series require a mapped workbook range.');
	const cells = new Map<
		string,
		{ sheet: number; row: number; col: number; value: string | number | null }
	>();
	for (const write of writes) {
		const { range, s: sheet } = host.target(write.formula);
		const horizontal = range.start.row === range.end.row;
		write.values.forEach((text, i) => {
			const row = range.start.row + (horizontal ? 0 : i),
				col = range.start.col + (horizontal ? i : 0);
			const cell = getCell(ctx.workbook.sheets[sheet]!, row, col);
			if (cell?.formula)
				throw new Error('Formula-dependent chart data must be edited in the worksheet.');
			const value = write.isNumeric ? (text === '' ? null : Number(text)) : text;
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new Error('Chart values must be finite.');
			if (cell?.richText && cell.value !== value)
				throw new Error('Rich chart source strings must be edited in the worksheet.');
			const key = `${sheet}!${formatCellAddress({ row, col })}`;
			const previous = cells.get(key);
			if (previous && previous.value !== value) throw new Error('Chart source writes conflict.');
			cells.set(key, { sheet, row, col, value });
		});
	}
	ctx.run(
		'Edit chart',
		'annotations',
		[{ kind: 'workbook' }],
		() => {
			for (const cell of cells.values()) {
				const sheet = ctx.workbook.sheets[cell.sheet]!;
				if (getCell(sheet, cell.row, cell.col)?.value === cell.value) continue;
				ensureCell(sheet, cell.row, cell.col);
				writeValue(sheet, cell.row, cell.col, cell.value);
				syncTableHeader(ctx, sheet, cell.row, cell.col);
			}
			chart.chartExSpace = next;
			const summary = chartExSummaryFromModel(next);
			if (summary.summary)
				chart.chartEx = {
					...summary.summary,
					kind: 'chart',
					anchor: chart.anchor,
					partName: chart.chartExPartName!,
				};
			else delete chart.chartEx;
		},
		{ sheet: s },
	);
}
