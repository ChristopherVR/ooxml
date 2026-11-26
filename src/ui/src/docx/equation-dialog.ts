import type { EditorView } from 'prosemirror-view';
import { dialogButton } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
} from './dialog-shell';

export function createEquationDialog(
	getView: () => EditorView | undefined,
	reportError?: (error: Error) => void,
) {
	const element = createDialogShell('Equation', 'dve-format-dialog dve-equation-dialog');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = dialogActions(cancel, ok);
	let locale: EditorLocale = 'en';
	let request = 0;
	let pending = false;
	let submit: (() => void) | undefined;
	const hide = () => {
		request++;
		pending = false;
		submit = undefined;
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	ok.addEventListener('click', () => submit?.());
	cancel.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	const dialog = {
		element,
		close: hide,
		setLocale(next: EditorLocale) {
			locale = next;
		},
		get isOpen() {
			return pending || isDialogOpen(element);
		},
		async open() {
			const view = getView();
			if (!view?.editable) return;
			const activeRequest = ++request;
			const selection = view.state.selection;
			pending = true;
			const loaded = await import('./equation-dialog-fields')
				.catch((cause: unknown) => {
					if (activeRequest === request)
						reportError?.(cause instanceof Error ? cause : new Error(String(cause)));
					return undefined;
				})
				.finally(() => {
					if (activeRequest === request) pending = false;
				});
			if (
				!loaded ||
				activeRequest !== request ||
				getView() !== view ||
				!view.editable ||
				!selection.eq(view.state.selection)
			)
				return;
			submit = loaded.showEquationFields({
				element,
				ok,
				actions,
				locale,
				view,
				selection,
				getView,
				hide,
			});
		},
	} satisfies FormatDialog;
	return dialog;
}
