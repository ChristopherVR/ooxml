import {
	formatAddress,
	formatRange,
	parseAddress,
	pivotAxisChoices,
	pivotRowFields,
	pivotSourceData,
	pivotTableAt,
	pivotValuesIndex,
	type PivotAggregate,
	type PivotLayout,
	type PivotTable,
	type PivotValuesPosition,
} from 'ooxml-core/xlsx';
import { target, type EditorContext } from 'ooxml-core/xlsx/ui';
import { checkbox, field, select, text, textInput } from './fields';
import { showDialog } from './frame';
import { pivotPageControls } from './pivot-page-filter';
import { pivotColumnControls, pivotRowControls } from './pivot-rows';
import { pivotValueControls, PIVOT_AGGREGATE_OPTIONS } from './pivot-values';

/** The Values placement choice of a report: 'first', 'last', 'between' or the fields before it. */
function placementValue(pivot: PivotTable): string {
	const fields = pivotRowFields(pivot).length;
	const at = pivotValuesIndex(pivot);
	return at < 0 || at >= fields
		? 'last'
		: at === 0
			? 'first'
			: at === 1 && fields === 2
				? 'between'
				: String(at);
}

const positionOf = (value: string): PivotValuesPosition =>
	value === 'first' || value === 'last' || value === 'between' ? value : Number(value);

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
	const more = pivotRowControls(
		ctx,
		pivot?.moreRowFields ?? [],
		() => updateRows(),
		pivot ? pivotAxisChoices(pivot) : [],
	);
	const moreColumns = pivotColumnControls(
		ctx,
		pivot?.moreColumnFields ?? [],
		() => updateRows(),
		pivot ? pivotAxisChoices(pivot) : [],
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
	const multipleNotice = text(ctx, 'Values are listed after the row fields.');
	// Where the measures sit among the row fields: one choice before, after and between them.
	const valuesPlacement = select(ctx, [['last', 'After row fields']], 'last');
	let placementChoices = '';
	let placementWanted = pivot ? placementValue(pivot) : 'last';
	const placeValues = (names: readonly string[]) => {
		const signature = names.join('\u0001');
		if (signature === placementChoices) return;
		placementChoices = signature;
		const options: (readonly [string, string])[] = [
			['first', ctx.t('Before row fields')],
			...names
				.slice(0, -1)
				.map((name, i): readonly [string, string] =>
					names.length === 2
						? ['between', ctx.t('Between row fields')]
						: [String(i + 1), ctx.t('After {field}', { field: name })],
				),
			['last', ctx.t('After row fields')],
		];
		valuesPlacement.replaceChildren(
			...options.map(([value, text]) => {
				const option = ctx.host.ownerDocument.createElement('option');
				option.value = value;
				option.textContent = text;
				return option;
			}),
		);
		valuesPlacement.value = options.some(([value]) => value === placementWanted)
			? placementWanted
			: 'last';
	};
	valuesPlacement.addEventListener('change', () => (placementWanted = valuesPlacement.value));
	const extras = pivotValueControls(
		ctx,
		pivot?.additionalValues ?? [],
		() => {
			updateRows();
		},
		pivot?.fields,
	);
	let pageData: Pick<import('ooxml-core/xlsx').PivotTable, 'fields' | 'records'> = pivot ?? {
		fields: [],
		records: [],
	};
	const page = pivotPageControls(ctx, pivot?.pageFilter, () => pageData);
	const loadFields = () => {
		let fields: string[];
		try {
			const data = pivotSourceData(ctx.workbook()!, t.sheet, source.value);
			fields = data.fields;
			pageData = pivot ?? data;
		} catch {
			fields = [];
			pageData = pivot ?? { fields: [], records: [] };
		}
		extras.load(fields);
		// A report with date groups offers their fields on the axes instead of the date field.
		const axisFields = pivot
			? pivotAxisChoices({ ...pivot, fields, records: pageData.records })
			: fields;
		more.load(axisFields);
		moreColumns.load(axisFields);
		for (const [control, chosen, optional] of [
			[row, pivot?.rowField, false],
			[second, pivot?.secondRowField, true],
			[column, pivot?.columnField, true],
			[value, pivot?.valueField, false],
		] as const) {
			const choices = control === value ? fields : axisFields;
			const previous = chosen ?? control.value;
			control.replaceChildren();
			for (const name of [...(optional ? [''] : []), ...choices]) {
				const option = ctx.host.ownerDocument.createElement('option');
				option.value = name;
				option.textContent = name || ctx.t('(None)');
				control.append(option);
			}
			if (choices.includes(previous ?? '') || (optional && !previous))
				control.value = previous ?? '';
			else if (control === value) control.value = fields.at(-1) ?? '';
		}
	};
	source.addEventListener('change', loadFields);
	loadFields();
	const updateRows = () => {
		extras.load(pageData.fields.filter((name) => name !== page.source.value));
		axis.disabled = extras.count() === 0;
		if (axis.disabled) axis.value = 'columns';
		more.enable(!!second.value);
		moreColumns.enable(!!column.value);
		page.update([
			row.value,
			second.value,
			...more.read(),
			column.value,
			...moreColumns.read(),
			value.value,
			...extras.read().map((v) => v.field),
		]);
		for (const option of second.options)
			option.disabled = !!page.source.value && option.value === page.source.value;
		const rowsAxis = axis.value === 'rows';
		const rowNames = [row.value, second.value, ...(second.value ? more.read() : [])].filter(
			Boolean,
		);
		placeValues(rowNames);
		valuesPlacement.disabled = !rowsAxis;
		if (valuesPlacement.disabled) valuesPlacement.value = 'last';
		const placed = valuesPlacement.value;
		multipleNotice.textContent =
			placed === 'first'
				? ctx.t('Values are listed before the row fields.')
				: placed === 'between'
					? ctx.t('Values are listed between the row fields.')
					: placed === 'last'
						? ctx.t('Values are listed after the row fields.')
						: ctx.t('Values are listed after {field}.', {
								field: rowNames[Number(placed) - 1] ?? '',
							});
		multipleNotice.hidden = !rowsAxis;
		// Subtotals are shown above a group only for fields that come after the Values.
		const before =
			placed === 'first'
				? 0
				: placed === 'last'
					? rowNames.length
					: placed === 'between'
						? 1
						: Number(placed);
		const valuesLate = rowsAxis && before > rowNames.length - 2;
		subtotals.input.disabled =
			!second.value || !!pivot?.fieldSubtotals?.some((entry) => entry.field === pivot.rowField);
		repeat.input.disabled = (!second.value && !rowsAxis) || layout.value === 'compact';
		position.disabled =
			!second.value || !subtotals.input.checked || layout.value === 'tabular' || valuesLate;
		tabularNotice.textContent = ctx.t(
			rowsAxis
				? 'Values on rows display subtotals below groups.'
				: 'Tabular layout displays subtotals below groups.',
		);
		tabularNotice.hidden =
			!second.value || !subtotals.input.checked || !(layout.value === 'tabular' || valuesLate);
	};
	axis.addEventListener('change', updateRows);
	valuesPlacement.addEventListener('change', updateRows);
	layout.addEventListener('change', updateRows);
	subtotals.input.addEventListener('change', updateRows);
	for (const control of [page.source, row, column, value, extras.host])
		control.addEventListener('change', updateRows);
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
			more.host,
			field(ctx, 'Report Layout:', layout),
			subtotals.wrapper,
			field(ctx, 'Subtotal Position:', position),
			tabularNotice,
			repeat.wrapper,
			field(ctx, 'Column Field:', column),
			moreColumns.host,
			...page.body,
			field(ctx, 'Value Field:', value),
			field(ctx, 'Summarize Values By:', aggregate),
			extras.host,
			field(ctx, 'Values Axis:', axis),
			field(ctx, 'Values Placement:', valuesPlacement),
			multipleNotice,
		],
		opened: () => (settings ? row : source).focus(),
		submit: () => {
			const pageFilter = page.read();
			const options = {
				...(pageFilter ? { pageFilter } : {}),
				rowField: row.value,
				valuesAxis: axis.value as 'columns' | 'rows',
				valuesPosition: positionOf(valuesPlacement.value),
				layout: layout.value as PivotLayout,
				subtotalPosition: position.value as 'top' | 'bottom',
				secondRowField: second.value,
				moreRowFields: second.value ? more.read() : [],
				subtotals: subtotals.input.checked,
				repeatLabels: repeat.input.checked,
				columnField: column.value,
				moreColumnFields: column.value ? moreColumns.read() : [],
				valueField: value.value,
				aggregate: aggregate.value as PivotAggregate,
				additionalValues: extras.read(),
			};
			if (pivot) {
				if (
					(pivot.pageFilter?.field ?? '') !== (pageFilter?.field ?? '') ||
					JSON.stringify(pivot.pageFilter?.item) !== JSON.stringify(pageFilter?.item) ||
					pivot.rowField !== options.rowField ||
					(pivot.valuesAxis ?? 'columns') !== options.valuesAxis ||
					placementValue(pivot) !== valuesPlacement.value ||
					(pivot.layout ?? 'tabular') !== options.layout ||
					(pivot.subtotalPosition ?? 'bottom') !== options.subtotalPosition ||
					(pivot.secondRowField ?? '') !== options.secondRowField ||
					JSON.stringify(pivot.moreRowFields ?? []) !== JSON.stringify(options.moreRowFields) ||
					(pivot.subtotals !== false) !== options.subtotals ||
					(pivot.repeatLabels ?? false) !== options.repeatLabels ||
					(pivot.columnField ?? '') !== options.columnField ||
					JSON.stringify(pivot.moreColumnFields ?? []) !==
						JSON.stringify(options.moreColumnFields) ||
					pivot.valueField !== options.valueField ||
					pivot.aggregate !== options.aggregate ||
					JSON.stringify(pivot.additionalValues ?? []) !== JSON.stringify(options.additionalValues)
				)
					t.session.refreshPivotTable(t.sheet, pivot.name, { ...options, pageFilter });
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
