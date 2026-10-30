// Review > Restrict Editing: formatting restrictions, the one kind of editing the document allows,
// the exceptions (editable regions for everyone) and Start/Stop Enforcing Protection with an
// optional password, as in Word's Restrict Editing pane.
import {
	enforceProtection,
	verifyProtectionPassword,
	type DocumentProtection,
} from 'ooxml-core/docx';
import { editRestrictionKey, setEditableRegion } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, fieldset, labelled, selectOf, textInput } from './dialog-fields';
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
import { translate, type EditorLocale } from './localization';

export interface RestrictEditingHost {
	view(): EditorView | undefined;
	protection(): DocumentProtection | undefined;
	/** Stores the protection in the document (undefined stops it) and redraws what depends on it. */
	apply(protection: DocumentProtection | undefined): void;
}

const TYPES = [
	['readOnly', 'No changes (Read only)'],
	['trackedChanges', 'Tracked changes'],
	['comments', 'Comments'],
	['forms', 'Filling in forms'],
] as const;

const password = () => {
	const input = textInput();
	input.type = 'password';
	input.autocomplete = 'new-password';
	return input;
};

export function createRestrictEditingDialog(host: RestrictEditingHost): FormatDialog {
	const element = createDialogShell('Restrict editing', 'dve-format-dialog');
	let locale: EditorLocale = 'en';
	const messages = [document.createElement('p'), document.createElement('p')] as const;
	for (const message of messages) message.setAttribute('role', 'alert');
	const say = (key: Parameters<typeof translate>[1] | '') => {
		for (const message of messages) message.textContent = key ? translate(locale, key) : '';
	};
	const formatting = checkbox('Limit formatting to a selection of styles');
	const restrict = checkbox('Allow only this type of editing in the document');
	const type = selectOf(TYPES);
	type.setAttribute('aria-label', 'Allow only this type of editing in the document');
	const everyone = checkbox('Everyone');
	const first = password();
	const again = password();
	const secret = password();
	const start = dialogButton('Yes, start enforcing protection', true);
	const stop = dialogButton('Stop protection', true);
	const done = dialogButton('Close');
	const enforcedView = [
		fieldset('Editing restrictions', messages[0]),
		labelled('Password', secret),
		dialogActions(done, stop),
	];
	const settingsView = [
		fieldset('Formatting restrictions', formatting.wrapper),
		fieldset('Editing restrictions', restrict.wrapper, type),
		fieldset('Exceptions (optional)', everyone.wrapper),
		fieldset(
			'Start enforcement',
			labelled('Enter new password (optional)', first),
			labelled('Reenter password to confirm', again),
			messages[1],
		),
		dialogActions(done, start),
	];
	const close = () => {
		setDialogOpen(element, false);
		element.replaceChildren();
		focusView(host.view());
	};
	const enforced = () => {
		const protection = host.protection();
		return Boolean(protection?.enforced && protection.edit !== 'none');
	};
	const render = () => {
		say('');
		const protection = host.protection();
		const current = enforced();
		element.replaceChildren(...(current ? enforcedView : settingsView));
		if (current) {
			secret.value = '';
			secret.parentElement!.hidden = !protection?.password && !protection?.unverifiablePassword;
			say('The document is protected');
		} else {
			formatting.input.checked = Boolean(protection?.formatting);
			const edit = protection?.edit;
			restrict.input.checked = Boolean(edit && edit !== 'none');
			type.value = edit && edit !== 'none' ? edit : 'readOnly';
			type.disabled = !restrict.input.checked;
			first.value = again.value = '';
			everyone.input.checked = selectionIsEditable();
		}
		localizeDialog(element, locale);
	};
	/** Whether the whole selection lies in a region everyone may edit. */
	const selectionIsEditable = (): boolean => {
		const view = host.view();
		if (!view) return false;
		const { from, to } = view.state.selection;
		return (
			from < to &&
			(editRestrictionKey.getState(view.state) ?? []).some(
				(range) => range.group === 'everyone' && range.from <= from && range.to >= to,
			)
		);
	};
	restrict.input.addEventListener('change', () => {
		type.disabled = !restrict.input.checked;
	});
	everyone.input.addEventListener('change', () => {
		const view = host.view();
		if (!view) return;
		const result = setEditableRegion(view.state, everyone.input.checked);
		if ('tr' in result) {
			view.dispatch(result.tr);
			say('');
		} else {
			everyone.input.checked = selectionIsEditable();
			say(
				result.error === 'empty'
					? 'Select text first'
					: 'This exception continues into another paragraph; change it in Word',
			);
		}
	});
	start.addEventListener('click', () => {
		const edit = restrict.input.checked ? (type.value as DocumentProtection['edit']) : 'none';
		if (edit === 'none') return say('Choose the type of editing to allow');
		if (first.value !== again.value) return say('The passwords do not match');
		host.apply(
			enforceProtection(edit, {
				...(first.value && { password: first.value }),
				...(formatting.input.checked && { formatting: true }),
			}),
		);
		const view = host.view();
		// Tracked changes protection records every change from here on, as Word turns the setting on.
		if (edit === 'trackedChanges' && view && !view.state.doc.attrs.trackChanges)
			view.dispatch(view.state.tr.setDocAttribute('trackChanges', true));
		close();
	});
	stop.addEventListener('click', () => {
		const verdict = verifyProtectionPassword(host.protection(), secret.value);
		if (verdict === undefined)
			return say('This protection uses a password that cannot be checked here');
		if (!verdict) return say('Incorrect password');
		const protection = host.protection();
		// Stopping keeps what was chosen, as Word does, but no longer enforces it.
		host.apply(
			protection && {
				edit: protection.edit,
				enforced: false,
				...(protection.formatting && { formatting: true }),
			},
		);
		close();
	});
	done.addEventListener('click', close);
	onDialogDismiss(element, close);
	return {
		element,
		get isOpen() {
			return isDialogOpen(element);
		},
		open() {
			render();
			setDialogOpen(element, true);
		},
		close,
		setLocale(value) {
			locale = value;
			localizeDialog(element, locale);
		},
	};
}
