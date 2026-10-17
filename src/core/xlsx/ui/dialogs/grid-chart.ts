// A chart over typed grid data (rows of numbers and text), for products that have no spreadsheet:
// the rows go into a scratch workbook and the same buildChart, chartView and renderChartSvg that
// Excel's Insert Chart uses draw them. Nothing here touches the DOM.
import {
	type CellValue,
	type ChartObject,
	type ChartType,
	type Workbook,
	chartView,
	createEditSession,
	createWorkbook,
	renderChartSvg,
} from '../../index';
import { buildChart, evaluateRef } from './insert-chart-model';

/** The chart types Insert Chart offers, in Office's order, with their English names. */
export const INSERT_CHART_TYPES: ReadonlyArray<readonly [type: ChartType, label: string]> = [
	['column', 'Column'],
	['line', 'Line'],
	['pie', 'Pie'],
	['bar', 'Bar'],
	['area', 'Area'],
	['scatter', 'Scatter'],
	['doughnut', 'Doughnut'],
	['radar', 'Radar'],
];

/** The types whose bars can be stacked (Insert Chart shows a Grouping choice for these). */
export const GROUPED_CHART_TYPES: readonly ChartType[] = ['column', 'bar'];

/** Largest grid read from pasted text: far beyond a readable chart, small enough to stay cheap. */
export const MAX_GRID_ROWS = 100;
export const MAX_GRID_COLS = 26;

/**
 * Rows of cells from text: one row per line, cells split on tabs (or commas when the line has no
 * tab). Numbers become numbers, empty cells null, and short rows are padded to the widest.
 */
export function parseChartGrid(text: string): CellValue[][] {
	const rows = text
		.split(/\r?\n/)
		.filter((line) => line.trim() !== '')
		.slice(0, MAX_GRID_ROWS)
		.map((line) =>
			(line.includes('\t') ? line.split('\t') : line.split(','))
				.slice(0, MAX_GRID_COLS)
				.map((cell): CellValue => {
					const trimmed = cell.trim();
					if (trimmed === '') return null;
					const number = Number(trimmed);
					return Number.isFinite(number) ? number : trimmed;
				}),
		);
	const width = rows.reduce((most, row) => Math.max(most, row.length), 0);
	return rows.map((row) => [...row, ...Array<CellValue>(width - row.length).fill(null)]);
}

export interface GridChart {
	workbook: Workbook;
	chart: ChartObject;
}

/** The chart for `rows` (text in the first row names the series, in the first column the categories). */
export function gridChart(
	rows: readonly (readonly CellValue[])[],
	type: ChartType,
	options: { grouping?: ChartObject['grouping']; title?: string } = {},
): GridChart | undefined {
	const width = rows[0]?.length ?? 0;
	if (rows.length === 0 || width === 0) return undefined;
	const workbook = createWorkbook();
	createEditSession(workbook).setRangeValues(
		0,
		{ row: 0, col: 0 },
		rows.map((row) => [...row]),
	);
	const range = { start: { row: 0, col: 0 }, end: { row: rows.length - 1, col: width - 1 } };
	return { workbook, chart: buildChart(workbook, 0, range, type, options) };
}

/** `grid` as a standalone SVG document `width` by `height`; `bare` leaves out the title and legend. */
export function gridChartSvg(grid: GridChart, width: number, height: number, bare = false): string {
	const chart = bare ? { ...grid.chart, title: '', showLegend: false } : grid.chart;
	const model = chartView(grid.workbook, 0, chart, (ref) => evaluateRef(grid.workbook, 0, ref));
	return renderChartSvg(model, width, height);
}
