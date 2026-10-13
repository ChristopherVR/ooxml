import {
	formatAddress,
	formatRange,
	parseAddress,
	pivotSourceData,
	pivotTableAt,
	type PivotAggregate,
	type PivotLayout,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, select, text, textInput } from './fields';
import { showDialog } from './frame';
import { pivotValueControls, PIVOT_AGGREGATE_OPTIONS } from './pivot-values';

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
	const second = select(ctx, [['', '(None)']], '');
	const layout = select(
		ctx,
		[
			['compact', 'Compact Form'],
			['outline', 'Outline Form'],
			['tabular', 'Tabular Form'],
		],
		pivot ? (pivot.layout ?? 'tabular') : 'compact',
	);
	const position = select(
		ctx,
		[
			['top', 'Above each group'],
			['bottom', 'Below each group'],
		],
		pivot?.subtotalPosition ?? (pivot ? 'bottom' : 'top'),
	);
	const subtotals = checkbox(ctx, 'Show subtotals', pivot?.subtotals !== false);
	const tabularNotice = text(ctx, 'Tabular layout displays subtotals below groups.');
	const repeat = checkbox(ctx, 'Repeat row labels', pivot?.repeatLabels ?? false);
	const column = select(ctx, [['', '(None)']], '');
	const value = select(ctx, [], '');
	const aggregate = select(ctx, PIVOT_AGGREGATE_OPTIONS, pivot?.aggregate ?? 'sum');
	const axis = select(
		ctx,
		[
			['columns', 'Values on Columns'],
			['rows', 'Values on Rows'],
		],
		pivot?.valuesAxis ?? 'columns',
	);
	const multipleNotice = text(ctx, 'Values on rows require one row field.');
	const extras = pivotValueControls(
		ctx,
		pivot?.additionalValues ?? [],
		() => {
			updateRows();
		},
		pivot?.fields,
	);
	const loadFields = () => {
		let fields: string[];
		try {
			fields = pivotSourceData(ctx.workbook()!, t.sheet, source.value).fields;
		} catch {
			fields = [];
		}
		extras.load(fields);
		for (const [control, chosen, optional] of [
			[row, pivot?.rowField, false],
			[second, pivot?.secondRowField, true],
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
	const updateRows = () => {
		axis.disabled = extras.count() === 0;
		if (axis.disabled) axis.value = 'columns';
		axis.querySelector<HTMLOptionElement>('[value="rows"]')!.disabled = !!second.value;
		const rowsAxis = axis.value === 'rows';
		second.disabled = rowsAxis;
		multipleNotice.hidden = !rowsAxis;
		subtotals.input.disabled = !second.value;
		repeat.input.disabled = (!second.value && !rowsAxis) || layout.value === 'compact';
		position.disabled = !second.value || !subtotals.input.checked || layout.value === 'tabular';
		tabularNotice.hidden = !second.value || !subtotals.input.checked || layout.value !== 'tabular';
	};
	axis.addEventListener('change', updateRows);
	layout.addEventListener('change', updateRows);
	subtotals.input.addEventListener('change', updateRows);
	second.addEventListener('change', updateRows);
	source.addEventListener('change', updateRows);
	updateRows();
	return showDialog(ctx, {
		name: settings ? 'pivot-settings' : 'create-pivot-table',
		heading: settings ? 'PivotTable Settings' : 'Create PivotTable',
		body: [
			field(ctx, 'Source Range:', source),
			field(ctx, 'Destination Cell:', destination),
			field(ctx, 'Row Field:', row),
			field(ctx, 'Second Row Field:', second),
			field(ctx, 'Report Layout:', layout),
			subtotals.wrapper,
			field(ctx, 'Subtotal Position:', position),
			tabularNotice,
			repeat.wrapper,
			field(ctx, 'Column Field:', column),
			field(ctx, 'Value Field:', value),
			field(ctx, 'Summarize Values By:', aggregate),
			extras.host,
			field(ctx, 'Values Axis:', axis),
			multipleNotice,
		],
		opened: () => (settings ? row : source).focus(),
		submit: () => {
			const options = {
				rowField: row.value,
				valuesAxis: axis.value as 'columns' | 'rows',
				layout: layout.value as PivotLayout,
				subtotalPosition: position.value as 'top' | 'bottom',
				secondRowField: second.value,
				subtotals: subtotals.input.checked,
				repeatLabels: repeat.input.checked,
				columnField: column.value,
				valueField: value.value,
				aggregate: aggregate.value as PivotAggregate,
				additionalValues: extras.read(),
			};
			if (pivot) {
				if (
					pivot.rowField !== options.rowField ||
					(pivot.valuesAxis ?? 'columns') !== options.valuesAxis ||
					(pivot.layout ?? 'tabular') !== options.layout ||
					(pivot.subtotalPosition ?? 'bottom') !== options.subtotalPosition ||
					(pivot.secondRowField ?? '') !== options.secondRowField ||
					(pivot.subtotals !== false) !== options.subtotals ||
					(pivot.repeatLabels ?? false) !== options.repeatLabels ||
					(pivot.columnField ?? '') !== options.columnField ||
					pivot.valueField !== options.valueField ||
					pivot.aggregate !== options.aggregate ||
					JSON.stringify(pivot.additionalValues ?? []) !== JSON.stringify(options.additionalValues)
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
