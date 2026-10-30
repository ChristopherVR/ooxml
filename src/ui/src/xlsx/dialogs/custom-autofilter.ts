// Custom AutoFilter: one or two comparisons joined by And / Or (also the Equals, Begins With,
// Greater Than, Between and date presets of the drop-down's Text, Number and Date Filters).
import {
	customFilter,
	decodeCondition,
	displayText,
	getCell,
	type CustomOperator,
	type FilterColumnKind,
} from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { dateInputToSerial, serialToDateInput } from './filter-dates';
import { el, field, radios, select, text } from './fields';
import { showDialog } from './frame';

export interface CustomAutoFilterProps {
	/** Zero-based sheet column of the filter column. */
	col?: number;
	kind?: FilterColumnKind;
	/** Operator preselected in the first row (a drop-down preset). */
	operator?: CustomOperator;
	/** Between: the two rows are "is greater than or equal to" and "is less than or equal to". */
	between?: boolean;
}

const OPERATORS: ReadonlyArray<readonly [CustomOperator, string]> = [
	['equals', 'equals'],
	['notEquals', 'does not equal'],
	['greaterThan', 'is greater than'],
	['greaterThanOrEqual', 'is greater than or equal to'],
	['lessThan', 'is less than'],
	['lessThanOrEqual', 'is less than or equal to'],
	['beginsWith', 'begins with'],
	['notBeginsWith', 'does not begin with'],
	['endsWith', 'ends with'],
	['notEndsWith', 'does not end with'],
	['contains', 'contains'],
	['notContains', 'does not contain'],
];

const DATE_OPERATORS = new Set<CustomOperator>([
	'equals',
	'notEquals',
	'greaterThan',
	'greaterThanOrEqual',
	'lessThan',
	'lessThanOrEqual',
]);

export function openCustomAutoFilter(
	ctx: EditorContext,
	props: CustomAutoFilterProps,
): Promise<boolean | undefined> {
	const t = target(ctx);
	const filter = t?.ws.autoFilter;
	if (!t || !filter) return Promise.resolve(undefined);
	const col = props.col ?? t.active.col;
	const offset = col - filter.range.start.col;
	const column = filter.columns?.find((c) => c.offset === offset);
	const kind = props.kind ?? 'text';
	const head = displayText(t.workbook, getCell(t.ws, filter.range.start.row, col));
	const rows = (column?.custom?.conditions ?? []).map(decodeCondition);
	const first: { op: CustomOperator; value: string } = rows[0] ?? {
		op: props.between ? 'greaterThanOrEqual' : (props.operator ?? 'equals'),
		value: '',
	};
	const second: { op: CustomOperator; value: string } | undefined = props.between
		? { op: 'lessThanOrEqual', value: '' }
		: rows[1];
	const date = kind === 'date';
	const operators = date ? OPERATORS.filter(([op]) => DATE_OPERATORS.has(op)) : OPERATORS;

	const build = (initial: { op: CustomOperator; value: string } | undefined) => {
		const op = select(ctx, operators, initial?.op ?? 'equals');
		op.setAttribute('aria-label', ctx.t('Operator'));
		const value = el(ctx, 'input', 'xve-input');
		value.type = date ? 'date' : 'text';
		value.setAttribute('aria-label', ctx.t('Value'));
		value.value = initial ? (date ? serialToDateInput(initial.value) : initial.value) : '';
		return { op, value };
	};
	const a = build(first);
	const b = build(second);
	const join = radios(
		ctx,
		'Combine with',
		[
			['and', 'And'],
			['or', 'Or'],
		],
		props.between || column?.custom?.and ? 'and' : 'or',
	);
	const wrap = (parts: ReturnType<typeof build>) => {
		const row = el(ctx, 'div', 'xve-row');
		row.append(parts.op, parts.value);
		return row;
	};
	const hint = date
		? text(ctx, 'Pick a date for each comparison.')
		: text(
				ctx,
				'Use ? to represent any single character and * to represent any series of characters.',
			);

	return showDialog<boolean>(ctx, {
		name: 'custom-autofilter',
		heading: 'Custom AutoFilter',
		body: [
			text(ctx, 'Show rows where: {name}', 'xve-note', { name: head }),
			field(ctx, 'First condition', wrap(a)),
			join.element,
			field(ctx, 'Second condition', wrap(b)),
			hint,
		],
		opened: () => a.value.focus(),
		submit: () => {
			const entries: { op: CustomOperator; value: string }[] = [];
			for (const part of [a, b]) {
				const raw = part.value.value.trim();
				if (!raw && part === b) continue;
				if (date) {
					const serial = dateInputToSerial(part.value.value, t.workbook.date1904);
					if (serial === undefined) {
						ctx.toast(ctx.t('Pick a date for each comparison.'), 'warning');
						return undefined;
					}
					entries.push({ op: part.op.value as CustomOperator, value: String(serial) });
				} else entries.push({ op: part.op.value as CustomOperator, value: part.value.value });
			}
			t.session.setFilterCriteria(t.sheet, col, customFilter(entries, join.get() === 'and'));
			return true;
		},
	});
}
