// Formulas > Evaluate Formula and Error Checking. The reduction steps and the rules come from the
// core's `audit` area; these dialogs only show them.
import {
	adjacentIssue,
	type ErrorIssue,
	type ErrorRule,
	findErrors,
	formatAddress,
	FormulaStepper,
	ignoreError,
	parseNumberText,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { el, text } from './fields';
import { button, showDialog } from './frame';
import { selectOn } from './go-to';

const evalBox = (ctx: EditorContext): HTMLDivElement => {
	const box = el(ctx, 'div', 'xve-eval-box');
	box.style.cssText =
		'font-family:Consolas,monospace;border:1px solid var(--xve-input,var(--office-border,#c8c6c4));padding:6px 8px;min-height:3em;white-space:pre-wrap;word-break:break-all';
	return box;
};

/** Evaluate Formula on the active cell, or on `at` (Error Checking's Show Calculation Steps). */
export function openEvaluateFormula(
	ctx: EditorContext,
	at?: { sheet: number; row: number; col: number },
): Promise<undefined> {
	const session = ctx.session();
	const active = ctx.selection.get().active;
	const start = at ?? { sheet: ctx.activeSheet(), row: active.row, col: active.col };
	if (!session || !FormulaStepper.canEvaluate(session.workbook, start)) {
		ctx.toast(ctx.t('Evaluate Formula needs a cell that contains a formula.'), 'warning');
		return Promise.resolve(undefined);
	}
	const stepper = new FormulaStepper(session.workbook, session.calc, start);
	const reference = evalBox(ctx);
	const evaluation = evalBox(ctx);
	const stepIn = button(ctx, 'Step In');
	const stepOut = button(ctx, 'Step Out');
	const evaluate = button(ctx, 'Evaluate', true);
	let finished = false;
	const paint = (): void => {
		const view = stepper.view();
		finished = view.finished;
		reference.textContent = view.reference;
		evaluation.replaceChildren();
		const { underline } = view;
		const doc = ctx.host.ownerDocument;
		if (underline) {
			const u = doc.createElement('u');
			u.textContent = view.text.slice(underline.start, underline.end);
			evaluation.append(
				doc.createTextNode(view.text.slice(0, underline.start)),
				u,
				doc.createTextNode(view.text.slice(underline.end)),
			);
		} else evaluation.textContent = view.text;
		stepIn.disabled = !view.canStepIn;
		stepOut.disabled = !view.canStepOut;
		evaluate.textContent = ctx.t(finished ? 'Restart' : 'Evaluate');
	};
	stepIn.addEventListener('click', () => (stepper.stepIn(), paint()));
	stepOut.addEventListener('click', () => (stepper.stepOut(), paint()));
	evaluate.addEventListener('click', () => {
		if (finished) stepper.restart();
		else stepper.evaluate();
		paint();
	});
	paint();
	const labelled = (key: string, box: HTMLElement): HTMLElement => {
		const wrap = el(ctx, 'div');
		wrap.append(text(ctx, key, 'xve-field-label'), box);
		return wrap;
	};
	return showDialog<undefined>(ctx, {
		name: 'evaluate-formula',
		heading: 'Evaluate Formula',
		okLabel: null,
		buttons: [stepIn, stepOut, evaluate],
		body: [labelled('Reference:', reference), labelled('Evaluation:', evaluation)],
		opened: () => evaluate.focus(),
	});
}

const RULE_TEXT: Record<ErrorRule, readonly [title: string, body: string]> = {
	'error-value': ['Error in cell {cell}', 'The formula in this cell results in the error {code}.'],
	'number-as-text': [
		'Number stored as text',
		'The number in this cell is formatted as text or preceded by an apostrophe.',
	],
	'inconsistent-formula': [
		'Inconsistent formula',
		'The formula in this cell differs from the formulas in this area of the spreadsheet.',
	],
	'omits-adjacent-cells': [
		'Formula omits adjacent cells',
		'The formula in this cell refers to a range that has additional numbers adjacent to it.',
	],
	'unlocked-formula': [
		'Unlocked cell containing formula',
		'This cell contains a formula and is not locked, so it can still be changed when the sheet is protected.',
	],
};

/** Error Checking: walks the issues of the active sheet from the active cell. */
export function openErrorChecking(ctx: EditorContext): Promise<undefined> {
	const session = ctx.session();
	if (!session) return Promise.resolve(undefined);
	const workbook = session.workbook;
	const sheet = ctx.activeSheet();
	const active = ctx.selection.get().active;
	let issues = findErrors(workbook, sheet);
	let current = adjacentIssue(issues, { sheet, row: active.row, col: active.col }, 1, true);
	const complete = (): void =>
		ctx.toast(ctx.t('The error check is complete for the entire sheet.'), 'info');
	if (!current) {
		complete();
		return Promise.resolve(undefined);
	}
	const title = el(ctx, 'strong');
	const body = el(ctx, 'p', 'xve-note');
	const ignore = button(ctx, 'Ignore Error');
	const edit = button(ctx, 'Edit in Formula Bar');
	const steps = button(ctx, 'Show Calculation Steps...');
	const convert = button(ctx, 'Convert to Number');
	const prev = button(ctx, 'Previous');
	const next = button(ctx, 'Next');
	let close = (): void => undefined;
	const show = (issue: ErrorIssue): void => {
		current = issue;
		selectOn(ctx, issue.sheet, { start: issue, end: issue });
		const [head, detail] = RULE_TEXT[issue.rule];
		const vars = {
			cell: formatAddress({ row: issue.row, col: issue.col }),
			code: issue.code ?? '',
		};
		title.textContent = ctx.t(head, vars);
		body.textContent = ctx.t(detail, vars);
		steps.hidden = issue.rule !== 'error-value';
		convert.hidden = issue.rule !== 'number-as-text';
		convert.disabled = ctx.readOnly();
		prev.disabled = next.disabled = issues.length < 2;
	};
	const go = (direction: 1 | -1): void => {
		if (current) {
			const target = adjacentIssue(issues, current, direction);
			if (target) show(target);
		}
	};
	prev.addEventListener('click', () => go(-1));
	next.addEventListener('click', () => go(1));
	ignore.addEventListener('click', () => {
		if (!current) return;
		const done = current;
		ignoreError(workbook, done);
		issues = findErrors(workbook, sheet);
		const target = adjacentIssue(issues, done, 1);
		if (!target) {
			close();
			complete();
		} else show(target);
	});
	edit.addEventListener('click', () => {
		close();
		ctx.grid()?.beginEdit();
	});
	steps.addEventListener('click', () => {
		if (!current) return;
		const at = { sheet: current.sheet, row: current.row, col: current.col };
		close();
		void openEvaluateFormula(ctx, at);
	});
	convert.addEventListener('click', () => {
		if (!current) return;
		const cell = workbook.sheets[current.sheet]?.rows.get(current.row)?.get(current.col);
		const value = typeof cell?.value === 'string' ? parseNumberText(cell.value) : undefined;
		if (value === undefined) return;
		session.setCellInput(current.sheet, current.row, current.col, String(value));
		issues = findErrors(workbook, sheet);
		const target = adjacentIssue(issues, current, 1);
		if (!target) {
			close();
			complete();
		} else show(target);
	});
	return showDialog<undefined>(
		ctx,
		{
			name: 'error-checking',
			heading: 'Error Checking',
			okLabel: null,
			buttons: [prev, next, ignore, edit, steps, convert],
			body: [title, body],
			opened: () => show(current as ErrorIssue),
		},
		(open) => (close = () => open.close()),
	);
}
