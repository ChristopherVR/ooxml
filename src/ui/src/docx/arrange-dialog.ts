import type { EditorView } from 'prosemirror-view';
import { dialogButton, labelled, numberInput } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { EditorCore } from './editor-core';

export function arrangementDialog(
	core: EditorCore,
	view: EditorView,
	pos: number,
	kind: 'alt' | 'rotate',
	value: string,
	apply: (value: string) => void,
) {
	const root = core.element.shadowRoot;
	if (!root || root.querySelector('.dve-arrange-dialog')) return;
	const dialog = createDialogShell(
		kind === 'alt' ? 'Alt Text' : 'More Rotation Options',
		'dve-format-dialog dve-arrange-dialog',
	);
	const input = kind === 'alt' ? document.createElement('textarea') : numberInput(-360, 360, 1);
	if (input instanceof HTMLTextAreaElement) {
		input.rows = 5;
		input.maxLength = 4000;
	} else {
		input.required = true;
		input.step = 'any';
	}
	input.value = value;
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const model = view.state.doc;
	const close = () => {
		setDialogOpen(dialog, false);
		dialog.remove();
	};
	cancel.addEventListener('click', close);
	onDialogDismiss(dialog, close);
	ok.addEventListener('click', () => {
		if (!input.reportValidity()) return;
		if (
			!view.isDestroyed &&
			view.editable &&
			core.targetView() === view &&
			view.state.doc === model &&
			view.state.selection.from === pos
		)
			apply(input.value);
		close();
	});
	dialog.append(
		labelled(kind === 'alt' ? 'Description' : 'Rotation (degrees)', input),
		dialogActions(cancel, ok),
	);
	root.append(dialog);
	localizeDialog(dialog, core.locale);
	setDialogOpen(dialog, true);
	input.focus();
}
