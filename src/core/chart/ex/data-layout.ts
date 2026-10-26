// Allocate ChartEx source data on a fresh worksheet without overwriting any existing source cells.
import type { ChartExSpace } from './model';
import { chartExLevelValues } from '../edit/chartex-sheet';
import { formatCellAddress } from '../workbook/cellref';
import { quoteChartSheetName } from '../workbook/layout';

/** A fresh worksheet block and the model whose references point at it. Existing parts are untouched. */
export function chartExDataLayout(
	space: ChartExSpace,
	sheetName: string,
): { space: ChartExSpace; rows: (string | number | null)[][] } {
	const draft = structuredClone(space);
	const columns: (string | number | null)[][] = [];
	const ref = (col: number, last: number) =>
		`${quoteChartSheetName(sheetName)}!$${formatCellAddress({ col, row: 0 }).replace(/1$/, '')}$2:$${formatCellAddress({ col, row: 0 }).replace(/1$/, '')}$${last + 1}`;
	for (const data of [
		...draft.chartData.data,
		...draft.plotArea.series.flatMap((series) => (series.data ? [series.data] : [])),
	]) {
		for (const dimension of data.dimensions) {
			if (!dimension.levels.length)
				throw new Error('ChartEx structural editing requires resolved source values.');
			const start = columns.length;
			const ordered =
				dimension.kind === 'strDim' ? [...dimension.levels].reverse() : dimension.levels;
			let count = 0;
			for (const level of ordered) {
				const values = chartExLevelValues(level);
				count = Math.max(count, values.length);
				const data = values.map((value) => {
					if (dimension.kind !== 'numDim') return value;
					if (value === '') return null;
					const n = Number(value);
					if (!Number.isFinite(n)) throw new Error('Numeric chart data is invalid.');
					return n;
				});
				columns.push([level.name ?? `Data ${columns.length + 1}`, ...data]);
			}
			if (!count) throw new Error('ChartEx source has no data points.');
			const first = ref(start, count).split(':')[0]!;
			const last = ref(columns.length - 1, count).split(':')[1]!;
			dimension.formula = {
				formula: `${first}:${last}`,
				...(dimension.formula?.direction ? { direction: 'col' } : {}),
			};
		}
	}
	for (const series of draft.plotArea.series) {
		const col = columns.length;
		columns.push([series.tx?.value ?? series.tx?.rich?.text ?? '']);
		series.tx = {
			formula: {
				formula: `${quoteChartSheetName(sheetName)}!${formatCellAddress({ col, row: 0 }).replace(/^([A-Z]+)(\d+)$/u, '$$$1$$$2')}`,
			},
			value: series.tx?.value ?? series.tx?.rich?.text ?? '',
		};
	}
	const count = Math.max(...columns.map((column) => column.length), 1);
	return {
		space: draft,
		rows: Array.from({ length: count }, (_, row) => columns.map((column) => column[row] ?? null)),
	};
}
