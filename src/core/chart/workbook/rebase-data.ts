// Allocate all classic chart vectors on a fresh worksheet for structural edits. Existing workbook
// ranges remain untouched; the returned model references only the new worksheet's literal cells.
import type { ChartSpace } from '../model';
import type { ChartDataCache, ChartDataSource } from '../model-series';
import { chartSeriesEntries } from '../edit/draft';
import { formatCellAddress } from './cellref';
import { quoteChartSheetName } from './layout';

export function rebaseChartData(
	space: ChartSpace,
	sheetName: string,
): {
	chartSpace: ChartSpace;
	rows: (string | number | null)[][];
} {
	const next = structuredClone(space);
	const rows: (string | number | null)[][] = [[]];
	let column = 0;
	const ref = (col: number, row: number) =>
		formatCellAddress({ col, row }).replace(/^([A-Z]+)(\d+)$/u, '$$$1$$$2');
	const formula = (from: number, to: number, count: number, row = 1) =>
		`${quoteChartSheetName(sheetName)}!${ref(from, row)}:${ref(to, row + Math.max(1, count) - 1)}`;
	const vector = (cache: ChartDataCache, heading: string): { col: number; count: number } => {
		const col = column++;
		rows[0]![col] = heading;
		let count = cache.pointCount ?? 0;
		for (const [at, point] of cache.points.entries()) {
			const index = point.index ?? at;
			if (!Number.isSafeInteger(index) || index < 0 || index > 1048574)
				throw new RangeError('Invalid chart cache index');
			count = Math.max(count, index + 1);
			const text = point.value;
			const value = cache.type === 'number' && text.trim() !== '' ? Number(text) : text;
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new RangeError('Invalid chart number');
			(rows[index + 1] ??= [])[col] = text === '' ? null : value;
		}
		if (count > 1048575 || column > 16384)
			throw new RangeError('Chart data exceeds worksheet bounds');
		for (let i = 1; i <= Math.max(1, count); i++) (rows[i] ??= [])[col] ??= null;
		return { col, count };
	};
	const source = (data: ChartDataSource | undefined, heading: string) => {
		if (!data) return;
		if (data.formula?.includes('['))
			throw new RangeError('External chart references cannot be rebased');
		if (data.levels) {
			if (!data.levels.length) throw new RangeError('Chart hierarchy has no cached levels');
			const vectors = [...data.levels]
				.reverse()
				.map((cache, i) => vector(cache, `${heading} ${i + 1}`));
			const count = Math.max(data.levelPointCount ?? 0, ...vectors.map((item) => item.count));
			data.formula = formula(vectors[0]!.col, vectors.at(-1)!.col, count);
			data.kind = 'multiLvlStrRef';
		} else {
			if (!data.cache) throw new RangeError('Chart data has no cached values to allocate');
			const { col, count } = vector(data.cache, heading);
			data.formula = formula(col, col, count);
			data.kind = data.cache.type === 'number' ? 'numRef' : 'strRef';
		}
	};
	for (const [index, { series }] of chartSeriesEntries(next).entries()) {
		const name =
			series.tx?.text ??
			series.tx?.value ??
			series.tx?.reference?.cache?.points[0]?.value ??
			`Series ${index + 1}`;
		source(series.categories, 'Category');
		source(series.xValues, `${name} X`);
		const valuesColumn = column;
		source(series.values ?? series.yValues, name);
		if (column === valuesColumn) throw new RangeError('Chart series has no values');
		series.tx = {
			text: name,
			reference: {
				kind: 'strRef',
				formula: `${quoteChartSheetName(sheetName)}!${ref(valuesColumn, 0)}`,
				cache: { type: 'string', pointCount: 1, points: [{ index: 0, value: name }] },
			},
		};
		source(series.bubbleSizes, `${name} Size`);
	}
	for (const row of rows) for (let col = 0; col < column; col++) row[col] ??= null;
	return { chartSpace: next, rows };
}
