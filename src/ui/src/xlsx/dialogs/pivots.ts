import {
	formatAddress,
	formatRange,
	parseAddress,
	pivotSourceData,
	pivotTableAt,
	type PivotAggregate,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { field, select, textInput } from './fields';
import { showDialog } from './frame';

const AGGREGATES = [
	['sum', 'Sum'],
	['count', 'Count'],
	['average', 'Average'],
	['min', 'Min'],
	['max', 'Max'],
] as const;

export function openPivotTable(ctx: EditorContext, settings = false): Promise<boolean | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const pivot = settings ? pivotTableAt(t.ws, t.range.start) : undefined;
	if (settings && !pivot) return Promise.resolve(undefined);
	const source = textInput(ctx, pivot?.source ?? formatRange(t.range));
	const destination = textInput(
		ctx,
		formatAddress(pivot?.location.start ?? { row: t.range.start.row, col: t.range.end.col + 2 }),
	);
	source.disabled = settings;
	destination.disabled = settings;
	const row = select(ctx, [], '');
	const column = select(ctx, [['', '(None)']], '');
	const value = select(ctx, [], '');
	const aggregate = select(ctx, AGGREGATES, pivot?.aggregate ?? 'sum');
	const loadFields = () => {
		let fields: string[];
		try {
			fields = pivotSourceData(ctx.workbook()!, t.sheet, source.value).fields;
		} catch {
			fields = [];
		}
		for (const [control, chosen, optional] of [
			[row, pivot?.rowField, false],
			[column, pivot?.columnField, true],
			[value, pivot?.valueField, false],
		] as const) {
			const previous = chosen ?? control.value;
			control.replaceChildren();
			for (const name of [...(optional ? [''] : []), ...fields]) {
				const option = ctx.host.ownerDocument.createElement('option');
				option.value = name;
				option.textContent = name || ctx.t('(None)');
				control.append(option);
			}
			if (fields.includes(previous ?? '') || (optional && !previous))
				control.value = previous ?? '';
			else if (control === value) control.value = fields.at(-1) ?? '';
		}
	};
	source.addEventListener('change', loadFields);
	loadFields();
	return showDialog(ctx, {
		name: settings ? 'pivot-settings' : 'create-pivot-table',
		heading: settings ? 'PivotTable Settings' : 'Create PivotTable',
		body: [
			field(ctx, 'Source Range:', source),
			field(ctx, 'Destination Cell:', destination),
			field(ctx, 'Row Field:', row),
			field(ctx, 'Column Field:', column),
			field(ctx, 'Value Field:', value),
			field(ctx, 'Summarize Values By:', aggregate),
		],
		opened: () => (settings ? row : source).focus(),
		submit: () => {
			const options = {
				rowField: row.value,
				columnField: column.value,
				valueField: value.value,
				aggregate: aggregate.value as PivotAggregate,
			};
			if (pivot) {
				if (
					pivot.rowField !== options.rowField ||
					(pivot.columnField ?? '') !== options.columnField ||
					pivot.valueField !== options.valueField ||
					pivot.aggregate !== options.aggregate
				)
					t.session.refreshPivotTable(t.sheet, pivot.name, options);
			} else {
				const at = parseAddress(destination.value.trim().replaceAll('$', ''));
				if (!at) throw new Error('The reference is not valid.');
				t.session.createPivotTable(t.sheet, at, { ...options, source: source.value });
				ctx.selection.set({ ranges: [{ start: at, end: at }], active: at, anchor: at });
			}
			return true;
		},
	});
}
