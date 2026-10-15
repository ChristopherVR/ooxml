import {
	PIVOT_SUBTOTAL_ORDER,
	fieldSubtotalFunctions,
	pivotRowFields,
	pivotTableAt,
	type PivotAggregate,
	type PivotFieldSubtotals,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, fieldset, radios, select } from './fields';
import { showDialog } from './frame';

/** What the dialog calls each function (Excel names Count Numbers in full, not as "Count"). */
const FUNCTION_LABELS: Record<PivotAggregate, string> = {
	sum: 'Sum',
	count: 'Count',
	average: 'Average',
	max: 'Max',
	min: 'Min',
	product: 'Product',
	countNums: 'Count Numbers',
	stdDev: 'StdDev',
	stdDevp: 'StdDevp',
	var: 'Var',
	varp: 'Varp',
};

/**
 * Field Settings: the subtotals of one row field. Automatic shows a Total row (or the group's
 * value when subtotals sit above), None shows no subtotal, and Custom shows a row for each chosen
 * function, below the group.
 */
export function openPivotFieldSettings(ctx: EditorContext): Promise<boolean | undefined> {
	const t = target(ctx);
	const pivot = t ? pivotTableAt(t.ws, t.active) : undefined;
	if (!t || !pivot || pivot.valuesAxis === 'rows') return Promise.resolve(undefined);
	const names = pivotRowFields(pivot);
	if (names.length < 2) return Promise.resolve(undefined);
	const chosen = select(
		ctx,
		names.map((name) => [name, name] as const),
		names[0],
		true,
	);
	const mode = radios(
		ctx,
		'Subtotals',
		[
			['auto', 'Automatic'],
			['none', 'None'],
			['custom', 'Custom'],
		],
		'auto',
	);
	const boxes = PIVOT_SUBTOTAL_ORDER.map((fn) => ({
		fn,
		box: checkbox(ctx, FUNCTION_LABELS[fn], false),
	}));
	const list = fieldset(
		ctx,
		'Select one or more functions:',
		...boxes.map((entry) => entry.box.wrapper),
	);
	const refresh = () => {
		for (const entry of boxes) entry.box.input.disabled = mode.get() !== 'custom';
	};
	const load = () => {
		const current = fieldSubtotalFunctions(pivot, chosen.value);
		mode.set(current === 'auto' ? 'auto' : current.length ? 'custom' : 'none');
		for (const entry of boxes)
			entry.box.input.checked = current !== 'auto' && current.includes(entry.fn);
		refresh();
	};
	chosen.addEventListener('change', load);
	for (const input of mode.inputs) input.addEventListener('change', refresh);
	load();
	return showDialog(ctx, {
		name: 'pivot-field-settings',
		heading: 'Field Settings',
		body: [field(ctx, 'Field:', chosen), mode.element, list],
		opened: () => chosen.focus(),
		submit: () => {
			const functions = boxes.filter((entry) => entry.box.input.checked).map((entry) => entry.fn);
			if (mode.get() === 'custom' && !functions.length)
				throw new Error('Choose at least one function.');
			const others: PivotFieldSubtotals[] = (pivot.fieldSubtotals ?? []).filter(
				(entry) => entry.field !== chosen.value,
			);
			const next =
				mode.get() === 'auto'
					? others
					: [...others, { field: chosen.value, functions: mode.get() === 'none' ? [] : functions }];
			t.session.refreshPivotTable(t.sheet, pivot.name, {
				fieldSubtotals: next,
				// The outer field's old on/off switch must not contradict the choice made here.
				...(chosen.value === pivot.rowField ? { subtotals: true } : {}),
			});
			return true;
		},
	});
}
