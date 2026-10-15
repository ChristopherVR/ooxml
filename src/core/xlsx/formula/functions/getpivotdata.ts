// GETPIVOTDATA over the worksheet-range pivots in `src/core/xlsx/pivot`. The value is
// aggregated from the pivot's cached records, so it matches the pivot's own grid cells
// (item totals, subtotals and the grand total) without reading the output range back.
import type { CellValue, PivotTable } from '../../model';
import { pivotTableAt } from '../../pivot/aggregate';
import { pivotItemLabel, pivotRowFields } from '../../pivot/data';
import { pivotMeasureNumber, pivotMeasures } from '../../pivot/measures';
import { pivotPageProjection } from '../../pivot/page-filter';
import { ERR, fail, isError, RefValue, type Scalar } from '../values';
import { scalar, spec, str } from './helpers';

const same = (cached: CellValue, wanted: Scalar): boolean => {
	if (typeof wanted === 'number' && typeof cached === 'number') return cached === wanted;
	if (cached === null) return wanted === null || wanted === '';
	if (typeof cached === 'object') return false;
	return String(pivotItemLabel(cached)).toLowerCase() === String(wanted ?? '').toLowerCase();
};

export const GETPIVOTDATA = spec(
	'GETPIVOTDATA',
	'Lookup & Reference',
	'GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...)',
	'Data stored in a PivotTable report.',
	2,
	254,
	(args, ctx) => {
		const ref = args[1];
		const area = ref instanceof RefValue ? ref.first : undefined;
		if (!area) fail(ERR.REF);
		const sheet = ctx.workbook.sheets[area.sheet];
		const found = sheet ? pivotTableAt(sheet, area.range.start) : undefined;
		if (!found || (args.length - 2) % 2 !== 0) fail(ERR.REF);
		let pivot: PivotTable;
		try {
			pivot = pivotPageProjection(found);
		} catch {
			return fail(ERR.REF);
		}
		const wanted = str(args[0]).toLowerCase();
		const measure = pivotMeasures(pivot).find(
			(m) => m.caption.toLowerCase() === wanted || m.field.toLowerCase() === wanted,
		);
		if (!measure) fail(ERR.REF);
		const axes = [...pivotRowFields(pivot), pivot.columnField, pivot.pageFilter?.field];
		let records = pivot.records;
		for (let i = 2; i < args.length; i += 2) {
			const name = str(args[i]);
			const index = pivot.fields.findIndex((f) => f.toLowerCase() === name.toLowerCase());
			const field = pivot.fields[index];
			if (index < 0 || field === undefined || !axes.includes(field)) fail(ERR.REF);
			const item = scalar(args[i + 1]);
			if (isError(item)) return item;
			records = records.filter((record) => same(record[index] ?? null, item));
		}
		if (records.length === 0) fail(ERR.REF);
		const value = pivotMeasureNumber(pivot, records, measure);
		return value;
	},
	['any'],
);
