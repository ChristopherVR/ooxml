// Grid keyboard: Ready-mode navigation and selection keys on the focus sink, and the in-cell
// editor's keys (commit / cancel / newline / point mode / F2 / F4). Keys the grid does not own
// propagate to the shell's global shortcuts.
import type { NavigationKey } from '@christophervr/xlsx-core';
import type { CellEditor } from './cell-editor.js';
import { currentRegionOrAll } from './grid-commands.js';
import type { GridSelection } from './grid-selection.js';
import type { GridView } from './grid-view.js';
import { toggleAbsolute } from './formula-text.js';

export interface KeyHost {
	view: GridView;
	selection: GridSelection;
	editor: CellEditor;
	openCellMenu(): void;
	openValidationList(): boolean;
	clearMarquee(): boolean;
	prepareCopy(cut: boolean): void;
}

const ARROWS: Record<string, ['up' | 'down' | 'left' | 'right', NavigationKey]> = {
	ArrowUp: ['up', 'ctrlUp'],
	ArrowDown: ['down', 'ctrlDown'],
	ArrowLeft: ['left', 'ctrlLeft'],
	ArrowRight: ['right', 'ctrlRight'],
};

/** The navigation key for an arrow / Home / End / Page key event, or undefined. */
export function navigationKey(
	event: Pick<KeyboardEvent, 'key' | 'ctrlKey' | 'metaKey'>,
): NavigationKey | undefined {
	const ctrl = event.ctrlKey || event.metaKey;
	const arrow = ARROWS[event.key];
	if (arrow) return ctrl ? arrow[1] : arrow[0];
	if (event.key === 'Home') return ctrl ? 'ctrlHome' : 'home';
	if (event.key === 'End') return ctrl ? 'ctrlEnd' : undefined;
	if (event.key === 'PageUp') return 'pageUp';
	if (event.key === 'PageDown') return 'pageDown';
	return undefined;
}

const consume = (event: KeyboardEvent): void => {
	event.preventDefault();
	event.stopPropagation();
};

/** Ready mode (not editing). */
export function readyKeyDown(host: KeyHost, event: KeyboardEvent): void {
	const { view, selection, editor } = host;
	const ctx = view.ctx;
	const ctrl = event.ctrlKey || event.metaKey;
	if (event.isComposing) return;
	const nav = navigationKey(event);
	if (nav && !event.altKey) {
		consume(event);
		selection.move(nav, event.shiftKey);
		return;
	}
	if (
		event.key === 'Alt' ||
		event.key === 'Control' ||
		event.key === 'Shift' ||
		event.key === 'Meta'
	)
		return;
	if (event.altKey && event.key === 'ArrowDown') {
		if (host.openValidationList()) consume(event);
		return;
	}
	switch (event.key) {
		case 'Enter':
			if (ctrl || event.altKey) return;
			consume(event);
			selection.advance(event.shiftKey ? 'up' : 'down');
			return;
		case 'Tab':
			if (ctrl || event.altKey) return;
			consume(event);
			selection.advance(event.shiftKey ? 'left' : 'right');
			return;
		case 'F2':
			if (event.shiftKey || ctrl || event.altKey) return;
			consume(event);
			editor.begin({ mode: 'edit' });
			return;
		case 'Delete':
			if (ctrl || event.altKey) return;
			consume(event);
			void ctx.commands.run('edit.delete');
			return;
		case 'Backspace':
			if (ctrl || event.altKey || ctx.readOnly()) return;
			consume(event);
			editor.begin({ mode: 'enter', text: '' });
			return;
		case 'Escape':
			if (host.clearMarquee()) consume(event);
			return;
		case 'ContextMenu':
			consume(event);
			host.openCellMenu();
			return;
		case 'F10':
			if (event.shiftKey) {
				consume(event);
				host.openCellMenu();
			}
			return;
	}
	const key = event.key.toLowerCase();
	if (ctrl && !event.altKey) {
		if (key === 'a' && !event.shiftKey) {
			consume(event);
			const sheet = view.sheet();
			const range = sheet ? currentRegionOrAll(sheet, selection.get()) : undefined;
			if (range)
				selection.set({
					...selection.get(),
					anchor: { ...selection.get().active },
					ranges: [range],
				});
			else selection.all();
			return;
		}
		if (key === ' ' || event.code === 'Space') {
			consume(event);
			const r = selection.get().ranges[selection.get().ranges.length - 1];
			selection.cols(r?.start.col ?? 0, r?.end.col ?? 0);
			return;
		}
		if (key === 'd' || key === 'r') {
			consume(event);
			void ctx.commands.run(key === 'd' ? 'edit.fill-down' : 'edit.fill-right');
			return;
		}
		if (key === 'c' || key === 'x') {
			// Let the native copy/cut happen on the sink; the clipboard module fills its data.
			event.stopPropagation();
			host.prepareCopy(key === 'x');
			return;
		}
		if (key === 'v') {
			event.stopPropagation();
			return;
		}
		return;
	}
	if (event.shiftKey && (key === ' ' || event.code === 'Space') && !ctrl) {
		consume(event);
		const r = selection.get().ranges[selection.get().ranges.length - 1];
		selection.rows(r?.start.row ?? 0, r?.end.row ?? 0);
	}
}

/** Edit / Enter / Point mode keys of the in-cell editor. */
export function editorKeyDown(host: KeyHost, event: KeyboardEvent): void {
	const { editor } = host;
	if (event.isComposing) return;
	if (editor.assist.handleKey(event, editor.field)) {
		event.stopPropagation();
		return;
	}
	const state = editor.bridge.state();
	const ctrl = event.ctrlKey || event.metaKey;
	event.stopPropagation();
	if (ctrl && !event.altKey && (event.key === ';' || event.key === ':')) {
		event.preventDefault();
		const time = event.shiftKey || event.key === ':';
		void host.view.ctx.commands.run(time ? 'edit.insert-time' : 'edit.insert-date');
		return;
	}
	switch (event.key) {
		case 'Enter':
			event.preventDefault();
			if (event.altKey) {
				const input = editor.field.input;
				const start = input.selectionStart ?? state.caret;
				const end = input.selectionEnd ?? state.caret;
				const text = `${input.value.slice(0, start)}\n${input.value.slice(end)}`;
				editor.setText(text, start + 1);
				return;
			}
			if (ctrl) editor.commit('none', true);
			else editor.commit(event.shiftKey ? 'up' : 'down', false);
			return;
		case 'Tab':
			event.preventDefault();
			editor.commit(event.shiftKey ? 'left' : 'right', false);
			return;
		case 'Escape':
			event.preventDefault();
			editor.cancel();
			return;
		case 'F2':
			event.preventDefault();
			editor.point.reset();
			editor.setMode(state.mode === 'edit' ? 'enter' : 'edit');
			return;
		case 'F4': {
			event.preventDefault();
			const next = toggleAbsolute(state.text, editor.field.input.selectionEnd ?? state.caret);
			if (next) editor.setText(next.text, next.caret);
			return;
		}
	}
	const nav = navigationKey(event);
	if (!nav || state.mode === 'edit' || event.altKey) return;
	if (editor.point.move(nav, event.shiftKey)) {
		event.preventDefault();
		return;
	}
	if (state.mode === 'enter' && ARROWS[event.key]) {
		event.preventDefault();
		if (editor.commit('none', false)) host.selection.move(nav, false);
	}
}
