// Inline sheet rename: an input replaces the tab label. Enter or blur commits through the core's
// `validateSheetName` and `renameSheet`; an invalid name shows a toast and keeps editing; Escape cancels.
import { validateSheetName } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { h } from '../grid/dom.js';

export interface RenameHandle {
	input: HTMLInputElement;
	/** Commits; false when the name was rejected (the input stays open). */
	commit(): boolean;
	cancel(): void;
}

export function startRename(
	ctx: EditorContext,
	tab: HTMLElement,
	index: number,
	done: (renamed: boolean) => void,
): RenameHandle | undefined {
	const session = ctx.session();
	const sheet = session?.workbook.sheets[index];
	if (!session || !sheet || ctx.readOnly() || session.workbook.structureLocked) return undefined;
	const doc = tab.ownerDocument;
	const input = h(doc, 'input', 'xst-rename', {
		type: 'text',
		maxlength: '31',
		'aria-label': ctx.t('Sheet name'),
	});
	input.value = sheet.name;
	const label = tab.querySelector<HTMLElement>('.xst-label');
	if (label) label.hidden = true;
	tab.append(input);
	let finished = false;

	const finish = (renamed: boolean) => {
		if (finished) return;
		finished = true;
		input.remove();
		if (label) label.hidden = false;
		done(renamed);
	};
	const commit = (): boolean => {
		if (finished) return true;
		const name = input.value.trim();
		if (name === sheet.name) {
			finish(false);
			return true;
		}
		const problem = validateSheetName(session.workbook, name, index);
		if (problem) {
			ctx.toast(ctx.t(problem), 'warning');
			input.focus();
			input.select();
			return false;
		}
		session.renameSheet(index, name);
		finish(true);
		return true;
	};
	const cancel = () => finish(false);

	input.addEventListener('keydown', (event) => {
		event.stopPropagation();
		if (event.key === 'Enter') {
			event.preventDefault();
			commit();
		} else if (event.key === 'Escape') {
			event.preventDefault();
			cancel();
		}
	});
	input.addEventListener('blur', () => {
		// A rejected name on blur falls back to the old one rather than trapping focus.
		if (!finished && !commit()) cancel();
	});
	// Clicks inside the input must not activate or drag the tab.
	for (const type of ['click', 'dblclick', 'pointerdown'])
		input.addEventListener(type, (event) => event.stopPropagation());
	input.focus();
	input.select();
	return { input, commit, cancel };
}
