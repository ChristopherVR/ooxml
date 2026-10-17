// A chart over typed grid data (rows of numbers and text), for products that have no spreadsheet:
// the rows go into a scratch workbook and the same buildChart, chartView and renderChartSvg that
// Excel's Insert Chart uses draw them. Nothing here touches the DOM.
import {
	type CellValue,
	type ChartObject,
	type ChartType,
	type ChartViewModel,
	type Workbook,
	chartView,
	createEditSession,
	createWorkbook,
	formatAddress,
	getCell,
	loadXlsx,
	renderChartSvg,
	saveXlsx,
	usedRange,
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

/**
 * The chart's cells in an embedded workbook, as Excel shows an OLE chart: the object is only the
 * chart, so the visible range (`oleSize`) is exactly its anchor, 10 columns by 20 rows (640 x 400
 * pixels at Excel's default sizes) to the right of the data.
 */
const EMBED_ANCHOR = { fromCol: 4, toCol: 14, rows: 20 } as const;
export const EMBED_ASPECT =
	((EMBED_ANCHOR.toCol - EMBED_ANCHOR.fromCol) * 64) / (EMBED_ANCHOR.rows * 20);

/**
 * An .xlsx whose first sheet holds `rows` and one chart of `type` beside them, ready to embed as
 * an object in another document (Visio's Insert > Chart). Undefined when there is no data.
 */
export async function gridChartWorkbook(
	rows: readonly (readonly CellValue[])[],
	type: ChartType,
	options: { grouping?: ChartObject['grouping']; title?: string } = {},
): Promise<Uint8Array | undefined> {
	const grid = gridChart(rows, type, options);
	if (!grid) return undefined;
	const { workbook, chart } = grid;
	chart.anchor = {
		from: { row: 0, col: EMBED_ANCHOR.fromCol, colOffset: 0, rowOffset: 0 },
		to: { row: EMBED_ANCHOR.rows, col: EMBED_ANCHOR.toCol, colOffset: 0, rowOffset: 0 },
	};
	const { kind: _kind, ...model } = chart;
	createEditSession(workbook).addChart(0, model);
	workbook.oleSize = `${formatAddress({ row: 0, col: EMBED_ANCHOR.fromCol })}:${formatAddress({
		row: EMBED_ANCHOR.rows - 1,
		col: EMBED_ANCHOR.toCol - 1,
	})}`;
	return saveXlsx(workbook);
}

/** A chart found in an embedded workbook, with the data it draws and how to draw it. */
export interface EmbeddedChart {
	/** The first sheet's cells from A1, as typed (rows padded to the widest). */
	rows: CellValue[][];
	chart: ChartObject;
	/** Resolved series and axes; draw with `renderChartSvg(view, width, height)`. */
	view: ChartViewModel;
}

/** The first chart of the first sheet that has one, or undefined when there is none. */
export async function readEmbeddedChart(bytes: Uint8Array): Promise<EmbeddedChart | undefined> {
	let workbook: Workbook;
	try {
		workbook = await loadXlsx(bytes);
	} catch {
		return undefined;
	}
	for (const [index, sheet] of workbook.sheets.entries()) {
		const chart = sheet.drawings.find(
			(drawing): drawing is ChartObject => drawing.kind === 'chart',
		);
		if (!chart) continue;
		const used = usedRange(sheet);
		const height = Math.min(used ? used.end.row + 1 : 0, MAX_GRID_ROWS);
		const width = Math.min(used ? used.end.col + 1 : 0, MAX_GRID_COLS);
		const rows = Array.from({ length: height }, (_, row) =>
			Array.from({ length: width }, (_, col): CellValue => getCell(sheet, row, col)?.value ?? null),
		);
		return {
			rows,
			chart,
			view: chartView(workbook, index, chart, (ref) => evaluateRef(workbook, index, ref)),
		};
	}
	return undefined;
}

/** Cells as the text the chart picker edits: tab separated, numbers as typed. */
export function chartGridText(rows: readonly (readonly CellValue[])[]): string {
	return rows
		.map((row) => row.map((cell) => (cell === null ? '' : String(cell))).join('\t'))
		.join('\n');
}
