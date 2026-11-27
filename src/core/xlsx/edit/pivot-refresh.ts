import type { PivotTable, PivotTablePatch } from '../model';
import {
	pivotCaption,
	pivotColumnFields,
	pivotRowFields,
	pivotSourceData,
	validatePivotFields,
} from '../pivot/data';
import { normalizedPivotMeasures } from '../pivot/measures';
import {
	pivotDateFieldExclusions,
	pivotDateFieldsOf,
	pivotGroupableDateFields,
	roundDateRecords,
} from '../pivot/dates';
import { reconcileDateAxes } from '../pivot/date-axes';
import { normalizedFieldSubtotals } from '../pivot/subtotals';
import { pruneItemFilters } from '../pivot/item-filters';
import { markDirty, pivotHasSlicers } from '../slicer/book';
import { sheetAt, type EditContext } from './context';
import { checkName, pivotAt, writePivot } from './pivots';

export function refreshPivotTable(
	ctx: EditContext,
	s: number,
	name: string,
	patch: PivotTablePatch = {},
): void {
	rebuildPivotTable(ctx, s, name, patch, 'source');
}

export function updatePivotTable(
	ctx: EditContext,
	s: number,
	name: string,
	patch: Omit<PivotTablePatch, 'source'>,
): void {
	rebuildPivotTable(ctx, s, name, patch, 'cache');
}

function rebuildPivotTable(
	ctx: EditContext,
	s: number,
	name: string,
	patch: PivotTablePatch,
	data: 'source' | 'cache',
): void {
	const old = pivotAt(ctx, s, name);
	const sheetId = sheetAt(ctx.workbook, s).sheetId;
	const connected = pivotHasSlicers(ctx.workbook, sheetId, old.name);
	if ((patch.name !== undefined || patch.source !== undefined) && connected)
		throw new Error(
			'Remove connected slicers and timelines before renaming or changing the source.',
		);
	if (patch.name !== undefined) checkName(ctx, patch.name, old.name);
	const aggregate = patch.aggregate ?? old.aggregate;
	const valueField = patch.valueField ?? old.valueField;
	const next: PivotTable = {
		...old,
		...(data === 'source' ? pivotSourceData(ctx.workbook, s, patch.source ?? old.source) : {}),
		name: patch.name ?? old.name,
		location: patch.destination
			? { start: { ...patch.destination }, end: { ...patch.destination } }
			: old.location,
		rowField: patch.rowField ?? old.rowField,
		valueField,
		aggregate,
		valueCaption:
			aggregate === old.aggregate && valueField === old.valueField
				? old.valueCaption
				: pivotCaption(aggregate, valueField, old),
	};
	if (patch.valueNumberFormat !== undefined) {
		if (typeof patch.valueNumberFormat === 'string' && patch.valueNumberFormat)
			next.valueNumberFormat = patch.valueNumberFormat;
		else delete next.valueNumberFormat;
	}
	if (next.rowField !== old.rowField) delete next.condition;
	if (patch.condition !== undefined) {
		if (patch.condition === null) delete next.condition;
		else {
			if (connected)
				throw new Error('Remove connected slicers and timelines before setting a condition.');
			next.condition = { ...patch.condition };
		}
	}
	if (patch.display) {
		const { style, ...display } = patch.display;
		if (
			style?.name !== undefined &&
			style.name !== '' &&
			!/^PivotStyle(?:Light|Medium|Dark)(?:[1-9]|1\d|2[0-8])$/.test(style.name)
		)
			throw new Error('Choose a built-in PivotTable style.');
		next.display = {
			...old.display,
			...display,
			...(style ? { style: { ...old.display?.style, ...style } } : {}),
		};
	}
	if (patch.columnField !== undefined) {
		if (patch.columnField) next.columnField = patch.columnField;
		else delete next.columnField;
	}
	if (patch.secondRowField !== undefined) {
		if (patch.secondRowField) next.secondRowField = patch.secondRowField;
		else delete next.secondRowField;
	}
	if (patch.columnField === '' && patch.moreColumnFields === undefined)
		delete next.moreColumnFields;
	if (patch.moreColumnFields !== undefined) {
		if (patch.moreColumnFields.length) next.moreColumnFields = [...patch.moreColumnFields];
		else delete next.moreColumnFields;
	}
	if (patch.secondRowField === '' && patch.moreRowFields === undefined) delete next.moreRowFields;
	if (patch.moreRowFields !== undefined) {
		if (patch.moreRowFields.length) next.moreRowFields = [...patch.moreRowFields];
		else delete next.moreRowFields;
	}
	if (patch.fieldSubtotals !== undefined) {
		if (patch.fieldSubtotals.length)
			next.fieldSubtotals = normalizedFieldSubtotals(patch.fieldSubtotals);
		else delete next.fieldSubtotals;
	}
	if (patch.dateGroups !== undefined) {
		if (patch.dateGroups.length) next.dateGroups = structuredClone(patch.dateGroups);
		else delete next.dateGroups;
	}
	if (ctx.workbook.date1904) next.date1904 = true;
	else delete next.date1904;
	if (patch.dateGroups !== undefined || next.dateGroups) reconcileDateAxes(next, old);
	if (patch.subtotals !== undefined) next.subtotals = patch.subtotals;
	if (patch.repeatLabels !== undefined) next.repeatLabels = patch.repeatLabels;
	if (patch.layout !== undefined) next.layout = patch.layout;
	if (patch.subtotalPosition !== undefined) next.subtotalPosition = patch.subtotalPosition;
	if (patch.valuesAxis !== undefined) next.valuesAxis = patch.valuesAxis;
	if (patch.pageFilter !== undefined) {
		if (patch.pageFilter === null) delete next.pageFilter;
		else next.pageFilter = structuredClone(patch.pageFilter);
	}
	if (patch.fieldGroups !== undefined) {
		if (patch.fieldGroups.length) next.fieldGroups = structuredClone(patch.fieldGroups);
		else delete next.fieldGroups;
	}
	if (patch.valuesPosition !== undefined) next.valuesPosition = patch.valuesPosition;
	else if (patch.valuesAxis === 'columns' || patch.secondRowField === '')
		delete next.valuesPosition;
	if (patch.additionalValues !== undefined) {
		const measures = normalizedPivotMeasures(patch.additionalValues, old);
		if (measures.length) next.additionalValues = measures;
		else delete next.additionalValues;
	}
	if (patch.calculatedField !== undefined) {
		if (patch.calculatedField === null) delete next.calculatedField;
		else
			next.calculatedField = {
				...patch.calculatedField,
				caption: pivotCaption('sum', patch.calculatedField.name, next),
			};
	}
	const dateExclusions =
		data === 'cache' && next.dateFields ? pivotDateFieldExclusions(next) : undefined;
	const dates =
		data === 'source'
			? pivotDateFieldsOf(ctx.workbook, s, next)
			: (next.dateFields?.filter((field) => !dateExclusions?.has(field)) ?? []);
	if (data === 'cache' && patch.dateGroups?.length) {
		const added = patch.dateGroups
			.map(({ field }) => field)
			.filter((field) => !dates.includes(field));
		if (added.length) {
			const groupable = pivotGroupableDateFields(ctx.workbook, s, next);
			const newlyTyped = added.filter((field) => groupable.includes(field));
			dates.push(...newlyTyped);
			roundDateRecords(next, newlyTyped, ctx.workbook.date1904);
		}
	}
	if (dates.length) next.dateFields = dates;
	else delete next.dateFields;
	if (data === 'source') roundDateRecords(next, dates, ctx.workbook.date1904);
	if (next.dateFilter && !dates.includes(next.dateFilter.field)) delete next.dateFilter;
	if (patch.fieldSubtotals === undefined && next.fieldSubtotals) {
		// Moving a field off the axes drops the subtotal setting it had.
		const onAxes = [...pivotRowFields(next), ...pivotColumnFields(next)];
		const left = next.fieldSubtotals.filter((entry) => onAxes.includes(entry.field));
		if (left.length) next.fieldSubtotals = left;
		else delete next.fieldSubtotals;
	}
	if (patch.itemFilters !== undefined) next.itemFilters = structuredClone(patch.itemFilters);
	const kept = patch.itemFilters ?? pruneItemFilters(next);
	if (kept) next.itemFilters = kept;
	else delete next.itemFilters;
	if (next.valuesAxis === 'rows' || !next.secondRowField) {
		if (next.display?.blankRows || next.display?.collapsed)
			next.display = { ...next.display, blankRows: false, collapsed: false };
	}
	validatePivotFields(next);
	writePivot(
		ctx,
		s,
		next,
		old,
		data === 'cache' ? 'PivotTable Settings' : undefined,
		connected ? () => markDirty(ctx.workbook) : undefined,
		data === 'cache',
	);
}

export function refreshAllPivotTables(ctx: EditContext): void {
	for (const [sheet, ws] of ctx.workbook.sheets.entries())
		for (const pivot of ws.pivotTables ?? []) refreshPivotTable(ctx, sheet, pivot.name);
}
