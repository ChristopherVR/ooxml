import type { EditorView } from 'prosemirror-view';
import { setReviewRecordingPreferences, type ReviewRecordingPreferences } from 'ooxml-core/docx/ui';
import { checkbox, dialogButton, fieldset } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { FormatDialog } from './font-dialog';
import { focusView } from './focus-view';
import type { EditorLocale } from './localization';

/** Document recording preferences share the ordinary dialog and core history/collaboration path. */
export function createReviewOptionsDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell('Tracking options', 'dve-format-dialog');
	const controls = {
		trackFormatting: checkbox('Track formatting'),
		trackMoves: checkbox('Track moves'),
	};
	const dirty = new Set<keyof ReviewRecordingPreferences>();
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	let locale: EditorLocale = 'en';
	const close = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	const content = [
		fieldset('Tracking', controls.trackFormatting.wrapper, controls.trackMoves.wrapper),
		dialogActions(cancel, ok),
	];
	for (const name of ['trackFormatting', 'trackMoves'] as const)
		controls[name].input.addEventListener('change', () => dirty.add(name));
	cancel.addEventListener('click', close);
	onDialogDismiss(element, close);
	ok.addEventListener('click', () => {
		const view = getView();
		if (!view?.editable) return;
		const preferences = Object.fromEntries(
			[...dirty].map((name) => [name, controls[name].input.checked]),
		) as ReviewRecordingPreferences;
		if (setReviewRecordingPreferences(preferences)(view.state, view.dispatch, view)) close();
	});
	return {
		element,
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			const view = getView();
			if (!view?.editable) return;
			dirty.clear();
			for (const name of ['trackFormatting', 'trackMoves'] as const)
				controls[name].input.checked = view.state.doc.attrs[name] !== false;
			element.append(...content);
			localizeDialog(element, locale);
			setDialogOpen(element, true);
		},
		close,
		setLocale(value) {
			locale = value;
			localizeDialog(element, locale);
		},
	};
}
