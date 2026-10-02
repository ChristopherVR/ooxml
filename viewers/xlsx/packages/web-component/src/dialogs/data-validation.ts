// Data > Data Validation: Settings, Input Message and Error Alert tabs over the core
// DataValidation of the active cell; OK sets it on every selected range in one undo step.
import {
	type ConditionalOperator,
	type DataValidation,
	type ValidationType,
	formatValue,
	parseCellInput,
	validationAt,
} from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { OPERATORS } from './cf-common.js';
import { checkbox, field, invalid, panel, select, tabs, textArea, textInput } from './fields.js';
import { button, showDialog } from './frame.js';

const ALLOW: ReadonlyArray<readonly [ValidationType, string]> = [
	['none', 'Any value'],
	['whole', 'Whole number'],
	['decimal', 'Decimal'],
	['list', 'List'],
	['date', 'Date'],
	['time', 'Time'],
	['textLength', 'Text length'],
	['custom', 'Custom'],
];

const STYLES = [
	['stop', 'Stop'],
	['warning', 'Warning'],
	['information', 'Information'],
] as const;

/** A stored operand as the field shows it. */
function shown(type: ValidationType, formula: string | undefined): string {
	if (formula === undefined) return '';
	const n = Number(formula);
	if (formula.trim() !== '' && Number.isFinite(n)) {
		if (type === 'date') return formatValue(n, 'm/d/yyyy').text;
		if (type === 'time') return formatValue(n, 'h:mm:ss').text;
		return formula;
	}
	if (type === 'list') {
		const quoted = /^"(.*)"$/.exec(formula);
		return quoted ? (quoted[1] ?? '') : `=${formula}`;
	}
	return `=${formula}`;
}

/** A typed operand as the model stores it (dates and times become serials). */
function stored(type: ValidationType, text: string): string | undefined {
	const value = text.trim();
	if (!value) return undefined;
	if (value.startsWith('=')) return value.slice(1);
	if (type === 'list') return `"${value.replace(/"/g, '')}"`;
	if (type === 'custom') return value;
	const parsed = parseCellInput(value).value;
	return typeof parsed === 'number' ? String(parsed) : undefined;
}

export function dataValidationDialog(
	ctx: EditorContext,
): Promise<DataValidation | null | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const current = validationAt(t.workbook, t.sheet, t.active.row, t.active.col);
	const allow = select(ctx, ALLOW, current?.type ?? 'none');
	const operator = select(ctx, OPERATORS, current?.operator ?? 'between');
	const min = textInput(ctx);
	const max = textInput(ctx);
	const source = textInput(ctx);
	const formula = textInput(ctx);
	const ignoreBlank = checkbox(ctx, 'Ignore blank', current?.allowBlank ?? true);
	const dropDown = checkbox(ctx, 'In-cell dropdown', current?.showDropDown ?? true);
	const fOperator = field(ctx, 'Data:', operator);
	const fMin = field(ctx, 'Minimum:', min);
	const fMax = field(ctx, 'Maximum:', max);
	const fSource = field(ctx, 'Source:', source);
	const fFormula = field(ctx, 'Formula:', formula);
	const showInput = checkbox(
		ctx,
		'Show input message when cell is selected',
		current?.showInputMessage ?? true,
	);
	const promptTitle = textInput(ctx, current?.promptTitle ?? '');
	const prompt = textArea(ctx, current?.prompt ?? '', 4);
	const showError = checkbox(
		ctx,
		'Show error alert after invalid data is entered',
		current?.showErrorMessage ?? true,
	);
	const errorStyle = select(ctx, STYLES, current?.errorStyle ?? 'stop');
	const errorTitle = textInput(ctx, current?.errorTitle ?? '');
	const error = textArea(ctx, current?.error ?? '', 4);
	const load = (dv: DataValidation | undefined): void => {
		const type = dv?.type ?? 'none';
		allow.value = type;
		operator.value = dv?.operator ?? 'between';
		const first = shown(type, dv?.formula1);
		min.value = first;
		source.value = type === 'list' ? first : '';
		formula.value = type === 'custom' ? first : '';
		max.value = shown(type, dv?.formula2);
	};
	const sync = (): void => {
		const type = allow.value as ValidationType;
		const ranged =
			type === 'whole' ||
			type === 'decimal' ||
			type === 'date' ||
			type === 'time' ||
			type === 'textLength';
		const two = operator.value === 'between' || operator.value === 'notBetween';
		fOperator.hidden = !ranged;
		fMin.hidden = !ranged;
		fMax.hidden = !ranged || !two;
		const minLabel = fMin.querySelector('span');
		if (minLabel) minLabel.textContent = ctx.t(two ? 'Minimum:' : 'Value:');
		fSource.hidden = type !== 'list';
		dropDown.wrapper.hidden = type !== 'list';
		fFormula.hidden = type !== 'custom';
		ignoreBlank.input.disabled = type === 'none';
	};
	load(current);
	allow.addEventListener('change', sync);
	operator.addEventListener('change', sync);
	sync();
	const pages = tabs(ctx, [
		{
			id: 'settings',
			label: 'Settings',
			panel: panel(
				ctx,
				field(ctx, 'Allow:', allow),
				ignoreBlank.wrapper,
				fOperator,
				fMin,
				fMax,
				fSource,
				dropDown.wrapper,
				fFormula,
			),
		},
		{
			id: 'input',
			label: 'Input Message',
			panel: panel(
				ctx,
				showInput.wrapper,
				field(ctx, 'Title:', promptTitle),
				field(ctx, 'Input message:', prompt),
			),
		},
		{
			id: 'error',
			label: 'Error Alert',
			panel: panel(
				ctx,
				showError.wrapper,
				field(ctx, 'Style:', errorStyle),
				field(ctx, 'Title:', errorTitle),
				field(ctx, 'Error message:', error),
			),
		},
	]);
	const clear = button(ctx, 'Clear All');
	clear.addEventListener('click', () => {
		load(undefined);
		ignoreBlank.input.checked = true;
		dropDown.input.checked = true;
		showInput.input.checked = true;
		showError.input.checked = true;
		errorStyle.value = 'stop';
		for (const input of [promptTitle, prompt, errorTitle, error]) input.value = '';
		sync();
	});
	const read = (): DataValidation | null | undefined => {
		const type = allow.value as ValidationType;
		const hasMessages = promptTitle.value || prompt.value || errorTitle.value || error.value;
		if (type === 'none' && !hasMessages) return null;
		const dv: DataValidation = { ranges: [], type, allowBlank: ignoreBlank.input.checked };
		if (
			type === 'whole' ||
			type === 'decimal' ||
			type === 'date' ||
			type === 'time' ||
			type === 'textLength'
		) {
			const two = operator.value === 'between' || operator.value === 'notBetween';
			dv.operator = operator.value as ConditionalOperator;
			const f1 = stored(type, min.value);
			if (f1 === undefined) return invalid(ctx, min, 'Enter a valid value.');
			dv.formula1 = f1;
			if (two) {
				const f2 = stored(type, max.value);
				if (f2 === undefined) return invalid(ctx, max, 'Enter a valid value.');
				dv.formula2 = f2;
			}
		} else if (type === 'list') {
			const f1 = stored(type, source.value);
			if (f1 === undefined)
				return invalid(ctx, source, 'The Source currently evaluates to an error.');
			dv.formula1 = f1;
			dv.showDropDown = dropDown.input.checked;
		} else if (type === 'custom') {
			const f1 = stored(type, formula.value);
			if (f1 === undefined) return invalid(ctx, formula, 'Enter a formula.');
			dv.formula1 = f1;
		}
		dv.showInputMessage = showInput.input.checked;
		dv.showErrorMessage = showError.input.checked;
		dv.errorStyle = errorStyle.value as DataValidation['errorStyle'] & string;
		if (promptTitle.value) dv.promptTitle = promptTitle.value;
		if (prompt.value) dv.prompt = prompt.value;
		if (errorTitle.value) dv.errorTitle = errorTitle.value;
		if (error.value) dv.error = error.value;
		return dv;
	};
	return showDialog<DataValidation | null>(ctx, {
		name: 'data-validation',
		heading: 'Data Validation',
		wide: true,
		body: pages.element,
		buttons: [clear],
		opened: () => allow.focus(),
		submit: () => {
			const dv = read();
			if (dv === undefined) return undefined;
			t.session.batch('Data validation', () => {
				for (const range of t.ranges) t.session.setDataValidation(t.sheet, dv ?? undefined, range);
			});
			return dv;
		},
	});
}
