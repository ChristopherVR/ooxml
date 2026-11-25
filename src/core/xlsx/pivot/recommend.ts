import { getCell } from '../cells';
import {
	isCellError,
	type CellValue,
	type PivotTable,
	type PivotTableOptions,
	type Workbook,
} from '../model';
import { isDateFormat } from '../numfmt/index';
import { styleAt } from '../styles';
import { pivotGrid } from './aggregate';
import { authoredCaptions } from './captions';
import { pivotCaption, pivotValueKey } from './data';
import { reconcileDateAxes } from './date-axes';
import { pivotSourceRange } from './source';

const MAX_SAMPLE_ROWS = 256;
const MAX_SAMPLE_COLUMNS = 32;
const MAX_CHOICES = 4;

export type PivotRecommendation = {
	options: PivotTableOptions;
	preview: CellValue[][];
};

export function recommendPivots(
	workbook: Workbook,
	sheet: number,
	source: string,
	locale: string,
): PivotRecommendation[] {
	const data = pivotSourceRange(workbook, sheet, source);
	const ws = workbook.sheets[data.sheet]!;
	const { start, end } = data.range;
	const height = end.row - start.row;
	if (height <= 0 || (height + 1) * (end.col - start.col + 1) > 1_000_000) return [];
	const sampleSize = Math.min(height, MAX_SAMPLE_ROWS);
	const width = Math.min(end.col - start.col + 1, MAX_SAMPLE_COLUMNS);
	const fields: string[] = [];
	const names = new Set<string>();
	const stats = Array.from({ length: width }, () => ({
		text: 0,
		numbers: 0,
		dates: 0,
		keys: new Set<string>(),
	}));
	for (let i = 0; i < width; i++) {
		const value = getCell(ws, start.row, start.col + i)?.value;
		if (
			value == null ||
			isCellError(value) ||
			String(value).length > 256 ||
			!String(value).trim() ||
			names.has(String(value).toLowerCase())
		)
			return [];
		fields.push(String(value));
		names.add(String(value).toLowerCase());
	}
	const records = Array.from({ length: sampleSize }, (_, i) => {
		const row = start.row + 1 + Math.floor((i * (height - 1)) / Math.max(1, sampleSize - 1));
		return stats.map((stat, j) => {
			const cell = getCell(ws, row, start.col + j);
			const value = cell?.value ?? null;
			if (value == null || value === '' || isCellError(value)) return value;
			if (typeof value === 'string' && value.length > 256) return null;
			if (typeof value === 'number') {
				stat.numbers++;
				if (
					Number.isFinite(value) &&
					value >= 1 &&
					value < 2_958_465 &&
					isDateFormat(styleAt(workbook, cell?.styleId).numFmt ?? 'General')
				)
					stat.dates++;
			} else stat.text++;
			stat.keys.add(pivotValueKey(value));
			return value;
		});
	});
	const values = fields.filter((_, i) => {
		const stat = stats[i]!;
		return stat.numbers > 0 && stat.text === 0 && stat.dates === 0;
	});
	const rows = fields
		.filter((_, i) => {
			const stat = stats[i]!;
			return stat.text > 0 && stat.numbers === 0 && stat.keys.size <= Math.max(16, sampleSize / 2);
		})
		.sort((a, b) => stats[fields.indexOf(a)]!.keys.size - stats[fields.indexOf(b)]!.keys.size);
	const date = fields.find((_, i) => stats[i]!.dates === sampleSize);
	const out: PivotRecommendation[] = [];
	const add = (rowField: string, valueField: string, group = false) => {
		const options = {
			source: data.source,
			rowField,
			valueField,
			locale,
			...(group ? { dateGroups: [{ field: rowField, by: ['years', 'months'] }] } : {}),
		} satisfies PivotTableOptions;
		const captions = authoredCaptions(options);
		const at = { row: 0, col: 0 };
		const report: PivotTable = {
			...options,
			fields,
			records,
			name: 'Preview',
			aggregate: 'sum',
			valueCaption: pivotCaption('sum', valueField, { captions }),
			location: { start: at, end: at },
			...(captions ? { captions } : {}),
			...(group ? { dateFields: [rowField], date1904: workbook.date1904 } : {}),
		};
		if (group) reconcileDateAxes(report);
		try {
			out.push({ options, preview: pivotGrid(report).values.slice(0, 10) });
		} catch (error) {
			if (!group) throw error;
			// An invalid sampled date grouping must not hide the categorical recommendations.
		}
	};
	if (date && values[0]) add(date, values[0], true);
	for (const value of values)
		for (const row of rows) {
			add(row, value);
			if (out.length === MAX_CHOICES) return out;
		}
	return out;
}
