import {
	hiddenItemKeys,
	pivotChartTable,
	pivotColumnFields,
	pivotFilterItems,
	pivotPageItemLabel,
	pivotRowFields,
	pivotTableAt,
	pivotValueKey,
	type ChartObject,
	type PivotChartButton,
	type PivotTable,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select } from './fields';
import { showDialog } from './frame';
import { openPivotTable } from './pivots';

export interface PivotChartFieldProps {
	chart: ChartObject;
	role: PivotChartButton;
	field: string;
}

export function openPivotChartField(ctx: EditorContext, props: PivotChartFieldProps) {
	const t = target(ctx);
	const bound = t && pivotChartTable(t.workbook, props.chart);
	if (!bound) return;
	if (
		props.role === 'value' ||
		props.role === 'filter' ||
		!bound.pivot.fields.includes(props.field)
	)
		return openPivotTable(ctx, { kind: 'settings', binding: { ...bound, role: props.role } });
	return openItems(ctx, bound.sheet, bound.pivot, props.field);
}

export function openPivotFieldFilter(ctx: EditorContext) {
	const t = target(ctx);
	const pivot = t && pivotTableAt(t.ws, t.active);
	if (t && pivot) return openItems(ctx, t.sheet, pivot);
}

function openItems(ctx: EditorContext, sheet: number, pivot: PivotTable, initial?: string) {
	const t = target(ctx);
	if (!t) return;
	const fields = [...pivotRowFields(pivot), ...pivotColumnFields(pivot)].filter((name) =>
		pivot.fields.includes(name),
	);
	const choice = select(
		ctx,
		fields.map((name) => [name, name]),
		initial ?? fields[0],
		true,
	);
	const list = el(ctx, 'div');
	list.setAttribute('role', 'group');
	const describe = () => {
		list.setAttribute('aria-label', choice.value);
		const items = pivotFilterItems(pivot, choice.value);
		const hidden = hiddenItemKeys(pivot, choice.value);
		if (!items?.length) {
			list.textContent = ctx.t('This field has no supported items or has too many items to list.');
			return [];
		}
		const controls = items.map((value) => {
			const label = el(ctx, 'label', 'xve-check');
			const input = el(ctx, 'input');
			input.type = 'checkbox';
			input.checked = !hidden.has(pivotValueKey(value));
			const name =
				value === null || value === '' ? ctx.t('(blank)') : String(pivotPageItemLabel(value));
			label.append(input, name);
			return { label, input, value };
		});
		list.replaceChildren(...controls.map(({ label }) => label));
		return controls;
	};
	let items = describe();
	choice.addEventListener('change', () => {
		items = describe();
	});
	return showDialog(ctx, {
		name: 'pivot-field-filter',
		heading: 'Filter PivotTable Field',
		body: [field(ctx, 'Field:', choice), list],
		opened: () => choice.focus(),
		submit: () => {
			if (!items.length) return true;
			const hidden = items.filter(({ input }) => !input.checked).map(({ value }) => value);
			const filters = (pivot.itemFilters ?? []).filter((entry) => entry.field !== choice.value);
			if (hidden.length) filters.push({ field: choice.value, hidden });
			t.session.refreshPivotTable(sheet, pivot.name, { itemFilters: filters });
			return true;
		},
	});
}
