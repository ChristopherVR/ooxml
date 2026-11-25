import type { EditorView } from 'prosemirror-view';
import { dialogButton } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { translate, type EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
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
			const loaded = await Promise.all([
				import('../math/equation-fields'),
				import('./equation-commands'),
			])
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
			const [{ createEquationFields }, commands] = loaded;
			const selected = commands.selectedEquation(view.state);
			const fields = createEquationFields(element.ownerDocument, {
				labels: {
					equation: translate(locale, 'Equation'),
					source: translate(locale, 'LaTeX equation'),
					display: translate(locale, 'Display equation on its own line'),
					preview: translate(locale, 'Preview'),
					empty: translate(locale, 'Type a LaTeX equation to see a preview.'),
					invalid: translate(locale, 'This LaTeX could not be converted to an equation.'),
				},
				...(selected ? { latex: selected.latex } : {}),
				display: selected?.display ?? commands.aloneOnItsLine(view.state),
				allowInline: true,
				onChange: (value) => {
					ok.disabled = !value;
				},
				accept: () => ok.click(),
			});
			fields.source.classList.add('dve-equation-source');
			fields.preview.classList.add('dve-equation-preview');
			submit = () => {
				const value = fields.draft;
				if (!value || getView() !== view || !selection.eq(view.state.selection)) return;
				if (!selected || value.latex !== selected.latex || value.display !== selected.display) {
					if (selected) commands.updateEquation(view, selected.pos, value.latex, value.display);
					else commands.insertEquation(view, value.latex, value.display);
				}
				hide();
			};
			element.replaceChildren(fields.element, actions);
			localizeDialog(element, locale);
			setDialogOpen(element, true);
			fields.source.focus();
		},
	} satisfies FormatDialog;
	return dialog;
}
