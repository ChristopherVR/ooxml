// Builds a core ChartObject from a cell range (Excel's Insert Chart over the selection) and
// resolves chart references against live cells for the preview.
import {
	type CellRange,
	type CellValue,
	type ChartObject,
	type ChartSeries,
	type ChartType,
	type Workbook,
	type Worksheet,
	formatAddress,
	getCell,
	normalizeRange,
	parseRange,
	quoteSheetName,
} from '@christophervr/xlsx-core';

const absolute = (r: CellRange): string => {
	const a = (row: number, col: number) =>
		formatAddress({ row, col }).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');
	const start = a(r.start.row, r.start.col);
	const end = a(r.end.row, r.end.col);
	return start === end ? start : `${start}:${end}`;
};

const value = (ws: Worksheet, row: number, col: number): CellValue =>
	getCell(ws, row, col)?.value ?? null;
const isText = (v: CellValue): boolean => typeof v === 'string' && v !== '';

/** A chart over `range`: text in the first row names the series, text in the first column labels the categories. */
export function buildChart(
	workbook: Workbook,
	sheetIndex: number,
	range: CellRange,
	type: ChartType,
	options: { grouping?: ChartObject['grouping']; title?: string } = {},
): ChartObject {
	const ws = workbook.sheets[sheetIndex];
	if (!ws) throw new RangeError(`No sheet at index ${sheetIndex}`);
	const r = normalizeRange(range);
	const prefix = `${quoteSheetName(ws.name)}!`;
	const rows = r.end.row - r.start.row + 1;
	const cols = r.end.col - r.start.col + 1;
	let headerRow = false;
	for (let c = r.start.col; c <= r.end.col; c++)
		if (isText(value(ws, r.start.row, c))) headerRow = true;
	let headerCol = false;
	for (let row = r.start.row + (headerRow ? 1 : 0); row <= r.end.row; row++)
		if (isText(value(ws, row, r.start.col))) headerCol = true;
	if (rows === 1) headerRow = false;
	if (cols === 1) headerCol = false;
	const byRows = cols - (headerCol ? 1 : 0) > rows - (headerRow ? 1 : 0) && type !== 'scatter';
	const series: ChartSeries[] = [];
	const dataRow0 = r.start.row + (headerRow ? 1 : 0);
	const dataCol0 = r.start.col + (headerCol ? 1 : 0);
	const lane = (fixed: number, from: number, to: number, alongRows: boolean): CellRange =>
		alongRows
			? { start: { row: fixed, col: from }, end: { row: fixed, col: to } }
			: { start: { row: from, col: fixed }, end: { row: to, col: fixed } };
	const read = (span: CellRange): CellValue[] => {
		const out: CellValue[] = [];
		for (let row = span.start.row; row <= span.end.row; row++)
			for (let col = span.start.col; col <= span.end.col; col++) out.push(value(ws, row, col));
		return out;
	};
	const count = byRows ? r.end.row - dataRow0 + 1 : r.end.col - dataCol0 + 1;
	for (let i = 0; i < count; i++) {
		const fixed = byRows ? dataRow0 + i : dataCol0 + i;
		const values = byRows
			? lane(fixed, dataCol0, r.end.col, true)
			: lane(fixed, dataRow0, r.end.row, false);
		const s: ChartSeries = {
			valuesRef: prefix + absolute(values),
			values: read(values).map((v) => (typeof v === 'number' ? v : null)),
			categories: [],
		};
		const head = byRows
			? headerCol
				? { row: fixed, col: r.start.col }
				: undefined
			: headerRow
				? { row: r.start.row, col: fixed }
				: undefined;
		if (head) {
			s.nameRef = prefix + absolute({ start: head, end: head });
			const name = value(ws, head.row, head.col);
			if (name !== null) s.name = String(name);
		}
		const cats = byRows
			? headerRow
				? lane(r.start.row, dataCol0, r.end.col, true)
				: undefined
			: headerCol
				? lane(r.start.col, dataRow0, r.end.row, false)
				: undefined;
		if (cats) {
			s.categoriesRef = prefix + absolute(cats);
			s.categories = read(cats).map((v) =>
				typeof v === 'number' ? v : v === null ? '' : String(v),
			);
		}
		series.push(s);
	}
	const chart: ChartObject = {
		kind: 'chart',
		anchor: {
			from: { row: r.start.row, col: r.end.col + 2, rowOffset: 0, colOffset: 0 },
			to: { row: r.start.row + 15, col: r.end.col + 10, rowOffset: 0, colOffset: 0 },
		},
		chartType: type,
		series,
		showLegend: series.length > 1 || type === 'pie' || type === 'doughnut',
		legendPosition: 'b',
	};
	if (options.grouping && (type === 'column' || type === 'bar')) chart.grouping = options.grouping;
	if (options.title) chart.title = options.title;
	return chart;
}

/** Resolves `Sheet1!$B$2:$B$5` (or a bare range on `sheetIndex`) to the cells' values. */
export function evaluateRef(workbook: Workbook, sheetIndex: number, ref: string): CellValue[] {
	const bang = ref.lastIndexOf('!');
	let ws = workbook.sheets[sheetIndex];
	if (bang >= 0) {
		const name = ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'");
		ws = workbook.sheets.find((s) => s.name === name) ?? ws;
	}
	const range = parseRange(ref.slice(bang + 1).replace(/\$/g, ''));
	if (!ws || !range) return [];
	const out: CellValue[] = [];
	for (let row = range.start.row; row <= range.end.row; row++)
		for (let col = range.start.col; col <= range.end.col; col++) out.push(value(ws, row, col));
	return out;
}
