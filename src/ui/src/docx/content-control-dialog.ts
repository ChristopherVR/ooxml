import type { EditorView } from 'prosemirror-view';
import { controlAtSelection, updateControlProperties } from 'ooxml-core/docx/ui';
import type { ContentControl, ContentControlLock } from 'ooxml-core/docx';
import { checkbox, dialogButton, labelled, textInput } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	createDialogShell,
	dialogActions,
	isDialogOpen,
	setDialogOpen,
	onDialogDismiss,
	localizeDialog,
} from './dialog-shell';

const lockOf = (noDelete: boolean, noEdit: boolean): ContentControlLock | undefined =>
	noDelete && noEdit
		? 'sdtContentLocked'
		: noDelete
			? 'sdtLocked'
			: noEdit
				? 'contentLocked'
				: undefined;

/** List entries typed one per line keep the stored value of an entry whose text did not change. */
export function parseListEntries(
	text: string,
	existing: ContentControl['items'] = [],
): NonNullable<ContentControl['items']> {
	return text
		.split(/\r\n|\r|\n/)
		.map((line) => line.trim())
		.filter(Boolean)
		.map((display) => ({
			display,
			value: existing.find((item) => item.display === display)?.value ?? display,
		}));
}

/** Developer > Properties: the title, tag, locking and kind-specific settings of the selected control. */
export function createContentControlDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = createDialogShell(
		'Content control properties',
		'dve-format-dialog dve-content-control-dialog',
	);
	const title = textInput();
	const tag = textInput();
	const noDelete = checkbox('Content control cannot be deleted');
	const noEdit = checkbox('Contents cannot be edited');
	const entries = document.createElement('textarea');
	entries.rows = 5;
	const entriesLabel = labelled('Drop-down list entries (one per line)', entries);
	const format = textInput();
	const formatLabel = labelled('Date display format', format);
	const message = document.createElement('p');
	message.setAttribute('role', 'status');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const content = [
		labelled('Title', title),
		labelled('Tag', tag),
		noDelete.wrapper,
		noEdit.wrapper,
		entriesLabel,
		formatLabel,
		message,
		dialogActions(cancel, ok),
	];
	let locale: EditorLocale = 'en';
	let target: ContentControl | undefined;
	const hide = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(getView());
	};
	cancel.addEventListener('click', hide);
	onDialogDismiss(element, hide);
	ok.addEventListener('click', () => {
		const view = getView();
		const span = view && controlAtSelection(view.state);
		if (!view?.editable || !span || !target) return hide();
		const next: ContentControl = { ...span.control };
		const set = <K extends 'alias' | 'tag'>(key: K, value: string) => {
			if (value) next[key] = value;
			else delete next[key];
		};
		set('alias', title.value);
		set('tag', tag.value);
		const lock = lockOf(noDelete.input.checked, noEdit.input.checked);
		if (lock) next.lock = lock;
		else delete next.lock;
		if (next.kind === 'dropDownList' || next.kind === 'comboBox')
			next.items = parseListEntries(entries.value, span.control.items);
		if (next.kind === 'date' && format.value.trim()) next.dateFormat = format.value.trim();
		view.dispatch(updateControlProperties(view.state, span, next));
		hide();
	});
	return {
		element,
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			const view = getView();
			if (!view?.editable) return;
			const span = controlAtSelection(view.state);
			target = span?.control;
			element.replaceChildren(...content);
			localizeDialog(element, locale);
			message.textContent = target ? '' : 'Select a content control to edit its properties.';
			localizeElement(message, locale);
			ok.disabled = !target;
			title.value = target?.alias ?? '';
			tag.value = target?.tag ?? '';
			noDelete.input.checked = target?.lock === 'sdtLocked' || target?.lock === 'sdtContentLocked';
			noEdit.input.checked =
				target?.lock === 'contentLocked' || target?.lock === 'sdtContentLocked';
			const isList = target?.kind === 'dropDownList' || target?.kind === 'comboBox';
			entriesLabel.hidden = !isList;
			entries.value = (target?.items ?? []).map((item) => item.display).join('\n');
			formatLabel.hidden = target?.kind !== 'date';
			format.value = target?.dateFormat ?? '';
			setDialogOpen(element, true);
			title.focus();
		},
	};
}
