/**
 * The chart data sheet: the categories x series table behind every product's "Edit Data" view
 * (PowerPoint and Word open it for an embedded chart, Excel edits the cells themselves). This is
 * the format-neutral model of that table and the edits a person can make to it; a product maps
 * the edits onto its own chart model (pptx: `applyChartDataSheetEdits` in the pptx area), so
 * formatting that lives on a series or a point is kept by the product, not rebuilt from here.
 *
 * Grid coordinates used by the sheet UI: row 0 is the header row (series names), column 0 the
 * category column; row `r >= 1` is category `r - 1` and column `c >= 1` is series `c - 1`.
 *
 * @module chart/data-sheet
 */

/**
 * One value of a series: a number, or `null` for a blank cell (written as a missing cached point).
 * The one series-values shape of the chart core: the data sheet holds it and the `chart/edit`
 * series operations (`setChartSeriesValues`, `addChartSeries`...) take it.
 */
export type ChartSeriesValue = number | null;

/** One series column: its name and one value per category (`null` for a blank cell). */
export interface ChartDataSheetSeries {
	name: string;
	values: readonly ChartSeriesValue[];
}

/** The neutral chart data shape: category labels down the rows, series across the columns. */
export interface ChartDataSheet {
	categories: readonly string[];
	series: readonly ChartDataSheetSeries[];
}

/** One edit of the sheet. A product applies these in order to its own chart model. */
export type ChartDataSheetEdit =
	| { kind: 'value'; series: number; category: number; value: number }
	| { kind: 'series-name'; series: number; name: string }
	| { kind: 'category-label'; category: number; label: string }
	| { kind: 'add-series' }
	| { kind: 'remove-series'; series: number }
	| { kind: 'add-category' }
	| { kind: 'remove-category'; category: number };

/** The name a new series gets: `Series N`, PowerPoint's own naming. */
export function nextChartSeriesName(seriesCount: number): string {
	return `Series ${seriesCount + 1}`;
}

/** The label a new category gets: `Cat N`, as the PowerPoint editor has always named it. */
export function nextChartCategoryName(categoryCount: number): string {
	return `Cat ${categoryCount + 1}`;
}

function inRange(index: number, length: number): boolean {
	return Number.isInteger(index) && index >= 0 && index < length;
}

/**
 * Whether `edit` can be applied to `sheet`: indices in range, a finite value, and never removing
 * the last series or category (a chart with no series or no category has nothing to plot, and
 * the sheet would have no row or column to add data back into).
 */
export function canApplyChartDataSheetEdit(
	sheet: ChartDataSheet,
	edit: ChartDataSheetEdit,
): boolean {
	const series = sheet.series.length;
	const categories = sheet.categories.length;
	switch (edit.kind) {
		case 'value':
			return (
				Number.isFinite(edit.value) &&
				inRange(edit.series, series) &&
				inRange(edit.category, categories)
			);
		case 'series-name':
			return inRange(edit.series, series);
		case 'category-label':
			return inRange(edit.category, categories);
		case 'remove-series':
			return series > 1 && inRange(edit.series, series);
		case 'remove-category':
			return categories > 1 && inRange(edit.category, categories);
		default:
			return true;
	}
}

function padded(values: readonly (number | null)[], length: number): (number | null)[] {
	const next = values.slice(0, length);
	while (next.length < length) next.push(null);
	return next;
}

/** Applies one edit, or returns `null` when {@link canApplyChartDataSheetEdit} refuses it. */
export function applyChartDataSheetEdit(
	sheet: ChartDataSheet,
	edit: ChartDataSheetEdit,
): ChartDataSheet | null {
	if (!canApplyChartDataSheetEdit(sheet, edit)) return null;
	const count = sheet.categories.length;
	switch (edit.kind) {
		case 'value':
			return {
				...sheet,
				series: sheet.series.map((entry, index) => {
					if (index !== edit.series) return entry;
					const values = padded(entry.values, count);
					values[edit.category] = edit.value;
					return { ...entry, values };
				}),
			};
		case 'series-name':
			return {
				...sheet,
				series: sheet.series.map((entry, index) =>
					index === edit.series ? { ...entry, name: edit.name } : entry,
				),
			};
		case 'category-label':
			return {
				...sheet,
				categories: sheet.categories.map((label, index) =>
					index === edit.category ? edit.label : label,
				),
			};
		case 'add-series':
			return {
				...sheet,
				series: [
					...sheet.series,
					{ name: nextChartSeriesName(sheet.series.length), values: sheet.categories.map(() => 0) },
				],
			};
		case 'remove-series':
			return { ...sheet, series: sheet.series.filter((_, index) => index !== edit.series) };
		case 'add-category':
			return {
				categories: [...sheet.categories, nextChartCategoryName(count)],
				series: sheet.series.map((entry) => ({
					...entry,
					values: [...padded(entry.values, count), 0],
				})),
			};
		case 'remove-category':
			return {
				categories: sheet.categories.filter((_, index) => index !== edit.category),
				series: sheet.series.map((entry) => ({
					...entry,
					values: padded(entry.values, count).filter((_, index) => index !== edit.category),
				})),
			};
	}
}

/** Applies edits in order, skipping any the sheet refuses at that point. */
export function applyChartDataSheetEdits(
	sheet: ChartDataSheet,
	edits: readonly ChartDataSheetEdit[],
): ChartDataSheet {
	let next = sheet;
	for (const edit of edits) next = applyChartDataSheetEdit(next, edit) ?? next;
	return next;
}

/** A cell of the sheet in grid coordinates (see the module comment). */
export interface ChartDataSheetCell {
	row: number;
	col: number;
}

/**
 * The edit that writes `raw` into one grid cell: a series name in the header row, a category
 * label in the category column, a value elsewhere. `parseNumber` reads a value cell (pass a
 * locale-aware parser); a value that does not parse gives `null`, as does the corner cell.
 */
export function chartDataSheetCellEdit(
	cell: ChartDataSheetCell,
	raw: string,
	parseNumber: (raw: string) => number | null,
): ChartDataSheetEdit | null {
	const { row, col } = cell;
	if (row < 0 || col < 0 || (row === 0 && col === 0)) return null;
	if (row === 0) return { kind: 'series-name', series: col - 1, name: raw };
	if (col === 0) return { kind: 'category-label', category: row - 1, label: raw };
	const value = parseNumber(raw);
	return value === null ? null : { kind: 'value', series: col - 1, category: row - 1, value };
}

/**
 * The edits that paste a block of cells (rows of text, as `parseTsv` reads the clipboard) with
 * its top-left corner on `anchor`. Like the spreadsheet behind PowerPoint's chart, a block that
 * runs past the last series or category adds the series and categories it needs first; a value
 * cell that does not parse is skipped, and the corner cell is never written.
 */
export function chartDataSheetPasteEdits(
	sheet: ChartDataSheet,
	anchor: ChartDataSheetCell,
	block: readonly (readonly string[])[],
	parseNumber: (raw: string) => number | null,
): ChartDataSheetEdit[] {
	const width = block.reduce((max, row) => Math.max(max, row.length), 0);
	if (block.length === 0 || width === 0) return [];
	const edits: ChartDataSheetEdit[] = [];
	const lastSeries = anchor.col + width - 1;
	const lastCategory = anchor.row + block.length - 1;
	for (let n = sheet.series.length; n < lastSeries; n++) edits.push({ kind: 'add-series' });
	for (let n = sheet.categories.length; n < lastCategory; n++) edits.push({ kind: 'add-category' });
	block.forEach((cells, r) =>
		cells.forEach((raw, c) => {
			const edit = chartDataSheetCellEdit(
				{ row: anchor.row + r, col: anchor.col + c },
				raw,
				parseNumber,
			);
			if (edit) edits.push(edit);
		}),
	);
	return edits;
}
