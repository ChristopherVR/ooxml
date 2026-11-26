import { getCell } from '../../cells';
import { isErrorCode } from '../../model';
import { linkedDataField } from '../../rich-values';
import { ERR, err, isError, RefValue } from '../values';
import { spec } from './helpers';

export const FIELDVALUE = spec(
	'FIELDVALUE',
	'Lookup & Reference',
	'FIELDVALUE(value, field_name)',
	'Returns a cached field from a linked data type.',
	2,
	2,
	([ref, name], ctx) => {
		if (isError(name)) return name;
		if (isError(ref)) return ref;
		if (typeof name !== 'string' || !(ref instanceof RefValue) || !ref.isCell()) return ERR.FIELD;
		const area = ref.first;
		const sheet = area && ctx.workbook.sheets[area.sheet];
		if (!area || !sheet) return ERR.REF;
		const { row, col } = area.range.start;
		ctx.readCell(area.sheet, row, col);
		const cell = getCell(sheet, row, col);
		const field = cell && linkedDataField(ctx.workbook, cell, name);
		if (!field) return ERR.FIELD;
		const { type, value } = field;
		if (type === 's') return value;
		if (type === 'e') return isErrorCode(value) ? err(value) : ERR.VALUE;
		if (type === 'b') {
			if (value === '1' || value === 'true') return true;
			if (value === '0' || value === 'false') return false;
			return ERR.VALUE;
		}
		if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?$/.test(value)) return ERR.VALUE;
		const number = Number(value);
		return Number.isFinite(number) && (type !== 'i' || Number.isInteger(number))
			? number
			: ERR.VALUE;
	},
	['any', 'value'],
);
