// Home > Clipboard: Paste (with values / formulas / formatting / transpose / Paste Special), Cut,
// Copy and Format Painter. The system clipboard belongs to the grid's `edit.*` commands; these
// commands delegate to them when the grid registered them and also keep the core payload so the
// paste variants can use it.
import type { ClipboardPayload, PasteMode } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, target, viewing } from './util.js';

interface ClipState {
	payload?: ClipboardPayload;
	painter?: { payload: ClipboardPayload; sticky: boolean };
	painterTimer?: ReturnType<typeof setTimeout>;
}

const states = new WeakMap<EditorContext, ClipState>();
export const clipState = (ctx: EditorContext): ClipState => {
	let state = states.get(ctx);
	if (!state) {
		state = {};
		states.set(ctx, state);
	}
	return state;
};

const delegated = (ctx: EditorContext, id: string): boolean => !!ctx.commands.get(id);

function copy(ctx: EditorContext, cut: boolean): void {
	const t = target(ctx);
	if (!t) return;
	const payload = cut ? t.session.cut(t.sheet, t.range) : t.session.copy(t.sheet, t.range);
	clipState(ctx).payload = payload;
	const grid = cut ? 'edit.cut' : 'edit.copy';
	if (delegated(ctx, grid)) void ctx.commands.run(grid);
}

/** Pastes the kept payload with a mode; falls back to the grid's system-clipboard paste. */
export function pasteWith(ctx: EditorContext, mode: PasteMode): void {
	const t = target(ctx);
	if (!t) return;
	const state = clipState(ctx);
	if (!state.payload) {
		if (mode === 'all' && delegated(ctx, 'edit.paste')) void ctx.commands.run('edit.paste');
		else if (mode === 'values' && delegated(ctx, 'edit.paste-values'))
			void ctx.commands.run('edit.paste-values');
		else ctx.toast(ctx.t('Copy cells first, then paste them.'), 'info');
		return;
	}
	const range = t.session.paste(t.sheet, t.active, state.payload, mode);
	if (state.payload.cut) state.payload = { ...state.payload, cut: false };
	ctx.selection.set({ ranges: [range], anchor: range.start, active: range.start });
}

export function clipboardCommands(): Command[] {
	return [
		editing({
			id: 'home.paste',
			label: 'Paste',
			icon: icon('paste'),
			shortcut: 'Ctrl+V',
			lock: false,
			run: (ctx, arg) => {
				const mode = typeof arg === 'string' ? (arg as PasteMode) : 'all';
				if (mode === 'all' && !clipState(ctx).payload && delegated(ctx, 'edit.paste'))
					return void ctx.commands.run('edit.paste');
				pasteWith(ctx, mode);
			},
		}),
		editing({
			id: 'home.paste-values',
			label: 'Paste Values',
			icon: icon('paste'),
			lock: false,
			run: (ctx) => pasteWith(ctx, 'values'),
		}),
		editing({
			id: 'home.paste-formulas',
			label: 'Paste Formulas',
			icon: icon('paste'),
			lock: false,
			run: (ctx) => pasteWith(ctx, 'formulas'),
		}),
		editing({
			id: 'home.paste-formatting',
			label: 'Paste Formatting',
			icon: icon('formatPainter'),
			lock: false,
			run: (ctx) => pasteWith(ctx, 'formats'),
		}),
		editing({
			id: 'home.paste-transpose',
			label: 'Transpose',
			icon: icon('paste'),
			lock: false,
			run: (ctx) => pasteWith(ctx, 'transpose'),
		}),
		editing({
			id: 'home.paste-special',
			label: 'Paste Special...',
			icon: icon('paste'),
			shortcut: 'Ctrl+Alt+V',
			lock: false,
			run: (ctx) => void ctx.dialogs.open('paste-special'),
		}),
		editing({
			id: 'home.cut',
			label: 'Cut',
			icon: icon('cut'),
			shortcut: 'Ctrl+X',
			lock: false,
			run: (ctx) => copy(ctx, true),
		}),
		viewing({
			id: 'home.copy',
			label: 'Copy',
			icon: icon('copy'),
			shortcut: 'Ctrl+C',
			run: (ctx) => copy(ctx, false),
		}),
		editing({
			id: 'home.format-painter',
			label: 'Format Painter',
			icon: icon('formatPainter'),
			lock: 'formatCells',
			checked: (ctx) => !!clipState(ctx).painter,
			run: (ctx, arg) => {
				const state = clipState(ctx);
				if (state.painter) {
					delete state.painter;
					return;
				}
				const t = target(ctx);
				if (!t) return;
				state.painter = { payload: t.session.copy(t.sheet, t.range), sticky: arg === 'sticky' };
			},
		}),
	];
}

/**
 * Applies an armed format painter to the next selection (after the drag settles), as Excel does
 * on mouse-up. Returns the unsubscribe function.
 */
export function installFormatPainter(ctx: EditorContext): () => void {
	return ctx.selection.onChange(() => {
		const state = clipState(ctx);
		if (!state.painter) return;
		if (state.painterTimer) clearTimeout(state.painterTimer);
		state.painterTimer = setTimeout(() => {
			const painter = state.painter;
			const t = target(ctx);
			if (!painter || !t || ctx.readOnly()) return;
			t.session.paste(t.sheet, t.range.start, painter.payload, 'formats');
			if (!painter.sticky) delete state.painter;
			ctx.requestRender();
		}, 250);
	});
}
