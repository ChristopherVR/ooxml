// The formula input of the formula bar: shows the active cell's content, edits it through the
// shared edit bridge (so the in-cell editor follows), offers function autocomplete and expands
// to several lines.
import type { EditorContext } from '../context.js';
import { h } from '../grid/dom.js';
import { editBridge, type EditState } from '../grid/edit-bridge.js';
import { createFormulaField, type FormulaField } from '../grid/formula-field.js';
import { createFunctionAssist, type FunctionAssist } from '../grid/function-assist.js';
import { splice } from '../grid/formula-text.js';
import { cellInputText } from './cell-input-text.js';

export interface FormulaInput {
	readonly field: FormulaField;
	readonly assist: FunctionAssist;
	/** Shows the active cell's content (when not editing). */
	refresh(): void;
	destroy(): void;
}

const ICON_X =
	'<svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M3.5 3.5l7 7M10.5 3.5l-7 7" stroke="currentColor" stroke-width="1.3"/></svg>';
const ICON_CHECK =
	'<svg width="14" height="14" viewBox="0 0 14 14" aria-hidden="true"><path d="M2.5 7.5l3 3 6-7" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>';
const ICON_CHEVRON =
	'<svg width="10" height="10" viewBox="0 0 10 10" aria-hidden="true"><path d="M2 3.5l3 3 3-3" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>';

function iconButton(
	doc: Document,
	className: string,
	label: string,
	svg: string,
): HTMLButtonElement {
	const button = h(doc, 'button', `xfb-btn ${className}`, {
		type: 'button',
		'aria-label': label,
		title: label,
	});
	button.innerHTML = svg;
	return button;
}

export function createFormulaInput(
	ctx: EditorContext,
	doc: Document,
	bar: HTMLElement,
): FormulaInput {
	const t = ctx.t;
	const bridge = editBridge(ctx);
	const actions = h(doc, 'div', 'xfb-actions');
	const cancel = iconButton(doc, 'xfb-cancel', t('Cancel'), ICON_X);
	const enter = iconButton(doc, 'xfb-enter', t('Enter'), ICON_CHECK);
	const fx = h(doc, 'button', 'xfb-btn xfb-fx', {
		type: 'button',
		'aria-label': t('Insert Function'),
		title: t('Insert Function'),
	});
	fx.innerHTML = '<i>fx</i>';
	actions.append(cancel, enter, fx);
	const field = createFormulaField(doc, 'xfb-input', t('Formula Bar'));
	const expand = h(doc, 'button', 'xfb-btn xfb-expand', {
		type: 'button',
		'aria-label': t('Expand Formula Bar'),
		title: t('Expand Formula Bar'),
		'aria-expanded': 'false',
	});
	expand.innerHTML = ICON_CHEVRON;
	bar.append(actions, field.root, expand);
	const assist = createFunctionAssist(doc, bar, t);
	const input = field.input;

	const showCell = () => {
		const wb = ctx.workbook();
		const { sheet, active } = ctx.selection.get();
		field.setText(wb ? cellInputText(wb, sheet, active.row, active.col) : '', 0);
	};
	const refresh = () => {
		const state = bridge.state();
		const ro = ctx.readOnly() || !ctx.session();
		input.readOnly = ro;
		fx.disabled = ro;
		bar.classList.toggle('xfb-editing', state.editing);
		if (!state.editing) showCell();
	};
	const anchor = () => ({
		left: field.root.offsetLeft,
		top: field.root.offsetTop + field.root.offsetHeight,
	});
	const updateAssist = () => assist.update(field, anchor());

	const onBridge = (state: Readonly<EditState>, origin: string) => {
		bar.classList.toggle('xfb-editing', state.editing);
		if (!state.editing) {
			assist.close();
			showCell();
			return;
		}
		if (origin !== 'bar') {
			field.setText(state.text, state.caret);
			assist.close();
		}
	};
	const begin = () => {
		if (bridge.state().editing || ctx.readOnly() || !ctx.session()) return;
		bridge.begin({
			mode: 'edit',
			source: 'bar',
			text: input.value,
			caret: input.selectionEnd ?? input.value.length,
		});
	};
	const onInput = () => {
		if (!bridge.state().editing) begin();
		if (!bridge.state().editing) return;
		bridge.update(input.value, input.selectionEnd ?? input.value.length, 'bar');
		updateAssist();
	};
	const onKey = (event: KeyboardEvent) => {
		event.stopPropagation();
		if (assist.handleKey(event, field)) return;
		const editing = bridge.state().editing;
		if (event.key === 'Escape') {
			event.preventDefault();
			if (editing) bridge.cancel();
			else showCell();
			ctx.grid()?.focus();
			return;
		}
		if (event.key === 'Enter' && event.altKey) {
			event.preventDefault();
			if (input.readOnly) return;
			const next = splice(input.value, input.selectionStart ?? 0, input.selectionEnd ?? 0, '\n');
			field.setText(next.text, next.caret);
			onInput();
			return;
		}
		if (!editing) return;
		if (event.key === 'Enter') {
			event.preventDefault();
			if (event.ctrlKey || event.metaKey) bridge.commit('none', true);
			else bridge.commit(event.shiftKey ? 'up' : 'down');
		} else if (event.key === 'Tab') {
			event.preventDefault();
			bridge.commit(event.shiftKey ? 'left' : 'right');
		}
	};
	const onCaret = () => {
		if (!bridge.state().editing) return;
		bridge.update(input.value, input.selectionEnd ?? input.value.length, 'bar');
		updateAssist();
	};
	const onFx = async () => {
		if (ctx.readOnly() || !ctx.session()) return;
		const { sheet, active } = ctx.selection.get();
		const name = await ctx.dialogs.open<unknown>('insert-function', { sheet, address: active });
		if (name !== undefined && typeof name !== 'string') return;
		const text = typeof name === 'string' && name ? `=${name.toUpperCase()}(` : '=';
		if (bridge.state().editing) bridge.update(text, text.length, 'bar');
		else bridge.begin({ text, mode: 'edit', source: 'bar', caret: text.length });
		field.setText(bridge.state().editing ? bridge.state().text : text);
		input.focus();
	};
	const onExpand = () => {
		const open = !bar.classList.contains('xfb-expanded');
		bar.classList.toggle('xfb-expanded', open);
		expand.setAttribute('aria-expanded', String(open));
		const label = open ? t('Collapse Formula Bar') : t('Expand Formula Bar');
		expand.setAttribute('aria-label', label);
		expand.title = label;
	};
	const keepFocus = (event: Event) => event.preventDefault();
	const onCancel = () => bridge.cancel();
	const onEnter = () => bridge.commit('none');
	const onFxClick = () => void onFx();

	input.addEventListener('focus', begin);
	input.addEventListener('mousedown', begin);
	input.addEventListener('input', onInput);
	input.addEventListener('keydown', onKey);
	input.addEventListener('keyup', onCaret);
	input.addEventListener('click', onCaret);
	input.addEventListener('blur', () => assist.close());
	for (const button of [cancel, enter, fx, expand]) button.addEventListener('mousedown', keepFocus);
	cancel.addEventListener('click', onCancel);
	enter.addEventListener('click', onEnter);
	fx.addEventListener('click', onFxClick);
	expand.addEventListener('click', onExpand);
	const offBridge = bridge.onChange(onBridge);
	refresh();

	return {
		field,
		assist,
		refresh,
		destroy() {
			offBridge();
			assist.destroy();
			actions.remove();
			field.root.remove();
			expand.remove();
		},
	};
}
