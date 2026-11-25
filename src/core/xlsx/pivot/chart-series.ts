import { columnLabel } from '../address';
import type { ChartSeries, PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotColumnFields, pivotRowFields, pivotValuesIndex } from './axes';
import { pivotItemLabeler } from './item-labels';
import { pivotMeasures } from './measures';

const MAX_SERIES = 128;
const MAX_POINTS = 50_000;
const MAX_LABEL = 256;

function chartLabel(values: Iterable<unknown>): string {
	let text = '';
	for (const value of values) {
		if (text) text += ' / ';
		text += String(value ?? '').slice(0, MAX_LABEL - text.length);
		if (text.length >= MAX_LABEL - 3) break;
	}
	return text;
}

/** A series corresponds to a Values field and one leaf combination of the column fields. */
export function pivotChartColumns(pivot: PivotTable, grid: PivotGrid) {
	const measures = pivotMeasures(pivot);
	const onRows = pivot.valuesAxis === 'rows';
	const columns = pivotColumnFields(pivot);
	const result = [];
	for (const [i, entry] of (grid.columnEntries ?? []).entries()) {
		if (entry.kind !== (columns.length ? 'leaf' : 'grand')) continue;
		for (const [measure, value] of measures.entries()) {
			const offset = (grid.rowLabelColumns ?? 1) + (onRows ? i : i * measures.length + measure);
			result.push({ entry, offset, measure, caption: value.caption });
			if (result.length === MAX_SERIES) return result;
		}
	}
	return result;
}

export function pivotChartSeries(pivot: PivotTable, grid: PivotGrid, sheet: string): ChartSeries[] {
	const label = pivotItemLabeler(pivot);
	const rowFields = pivotRowFields(pivot);
	const columnFields = pivotColumnFields(pivot);
	const valuesAt = pivotValuesIndex(pivot);
	const series = pivotChartColumns(pivot, grid);
	const limit = Math.floor(MAX_POINTS / Math.max(1, series.length));
	const rows = new Map<number, { row: number; category: string }[]>();
	for (const [i, entry] of (grid.rowEntries ?? []).entries()) {
		if (entry.type || entry.indices.length !== rowFields.length + Number(valuesAt >= 0)) continue;
		const measure = entry.measure ?? 0;
		let list = rows.get(measure);
		if (!list) rows.set(measure, (list = []));
		if (list.length >= limit) continue;
		const category = chartLabel(
			rowFields.map((field, level) => {
				const index = entry.indices[level + Number(valuesAt >= 0 && level >= valuesAt)];
				return label(field, grid.axisItems?.[level]?.[index ?? -1] ?? null);
			}),
		);
		list.push({ row: grid.firstDataRow + i, category });
	}
	const start = pivot.location.start;
	const cell = (row: number, col: number) =>
		`$${columnLabel(start.col + col)}$${start.row + row + 1}`;
	const prefix = `'${sheet.replaceAll("'", "''")}'!`;
	return series.map(({ entry, offset, measure, caption }) => {
		const path = entry.values.map((value, i) => label(columnFields[i] ?? '', value));
		const name = path.length
			? chartLabel([...path, ...(pivot.additionalValues?.length ? [caption] : [])])
			: caption.slice(0, MAX_LABEL);
		const data = rows.get(valuesAt < 0 ? 0 : measure) ?? [];
		const first = data[0]?.row;
		const last = data.at(-1)?.row;
		const ref = (col: number) =>
			first === undefined || last === undefined
				? undefined
				: `${prefix}${cell(first, col)}:${cell(last, col)}`;
		const categoriesRef = ref(0);
		const valuesRef = ref(offset);
		return {
			name,
			categories: data.map(({ category }) => category),
			values: data.map(({ row }) => {
				const value = grid.values[row]?.[offset];
				return typeof value === 'number' && Number.isFinite(value) ? value : null;
			}),
			...(categoriesRef ? { categoriesRef } : {}),
			...(valuesRef ? { valuesRef } : {}),
		};
	});
}
