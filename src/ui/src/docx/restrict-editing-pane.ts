// Review > Restrict Editing as Word's docked task pane: formatting restrictions, the one kind of
// editing the document allows, the exceptions (editable regions for everyone), Find next region I
// can edit, Highlight the regions I can edit, and Start/Stop Enforcing Protection with an
// optional password.
import {
	enforceProtection,
	verifyProtectionPassword,
	type DocumentProtection,
} from 'ooxml-core/docx';
import { editRestrictionKey, nextEditableRegion, setEditableRegion } from 'ooxml-core/docx/ui';
import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, fieldset, labelled, selectOf, textInput } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { focusView } from './focus-view';
import { localizeElement, translate, type EditorLocale } from './localization';

export interface RestrictEditingHost {
	view(): EditorView | undefined;
	protection(): DocumentProtection | undefined;
	/** Stores the protection in the document (undefined stops it) and redraws what depends on it. */
	apply(protection: DocumentProtection | undefined): void;
	/** The review author, whose exceptions count. */
	user(): string;
	highlight(): boolean;
	setHighlight(on: boolean): void;
}

const TYPES = [
	['readOnly', 'No changes (Read only)'],
	['trackedChanges', 'Tracked changes'],
	['comments', 'Comments'],
	['forms', 'Filling in forms'],
] as const;

const STYLE = `.dve-restrict-pane{flex:0 0 300px;min-width:0;max-width:40vw;display:flex;flex-direction:column;overflow:auto;background:var(--surface,#fff);border-left:1px solid var(--line,#ddd);font:13px/1.4 'Segoe UI',sans-serif;color:var(--ink,#222)}
.dve-restrict-pane[hidden]{display:none}.dve-restrict-head{display:flex;justify-content:space-between;align-items:center;padding:12px}.dve-restrict-head button{border:0;background:none;color:inherit;font-size:20px;cursor:pointer}.dve-restrict-body{padding:0 12px 12px;display:flex;flex-direction:column;gap:8px}.dve-restrict-body fieldset{border:0;border-top:1px solid var(--line,#ddd);margin:0;padding:8px 0;display:flex;flex-direction:column;gap:6px}.dve-restrict-body legend{font-weight:600;padding:0 4px 0 0}.dve-restrict-body label{display:flex;flex-direction:column;gap:2px}.dve-restrict-body label.dve-dialog-check{flex-direction:row;align-items:center;gap:6px}.dve-restrict-body p{margin:0;color:var(--muted,#666)}.dve-restrict-actions{display:flex;gap:8px;flex-wrap:wrap}`;

const password = () => {
	const input = textInput();
	input.type = 'password';
	input.autocomplete = 'new-password';
	return input;
};

export function createRestrictEditingPane(host: RestrictEditingHost): FormatDialog {
	const element = document.createElement('aside');
	element.className = 'dve-restrict-pane';
	element.setAttribute('aria-label', 'Restrict editing');
	element.dataset.title = 'Restrict editing';
	element.hidden = true;
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
	// The controls both views share are made twice: a node can only be in one place.
	const highlights: HTMLInputElement[] = [];
	const regionControls = () => {
		const highlight = checkbox('Highlight the regions I can edit');
		const find = dialogButton('Find next region I can edit');
		highlights.push(highlight.input);
		highlight.input.addEventListener('change', () => host.setHighlight(highlight.input.checked));
		find.addEventListener('click', () => {
			const view = host.view();
			const range = view && nextEditableRegion(view.state, host.user());
			if (!view || !range) return say('There are no regions you can edit');
			say('');
			view.dispatch(view.state.tr.setSelection(range).scrollIntoView());
			view.focus();
		});
		return [find, highlight.wrapper] as const;
	};
	const enforcedView = [
		fieldset('Editing restrictions', messages[0]),
		fieldset('Exceptions (optional)', ...regionControls()),
		labelled('Password', secret),
		stop,
	];
	const settingsView = [
		fieldset('Formatting restrictions', formatting.wrapper),
		fieldset('Editing restrictions', restrict.wrapper, type),
		fieldset('Exceptions (optional)', everyone.wrapper, ...regionControls()),
		fieldset(
			'Start enforcement',
			labelled('Enter new password (optional)', first),
			labelled('Reenter password to confirm', again),
			messages[1],
			start,
		),
	];
	const head = document.createElement('div');
	head.className = 'dve-restrict-head';
	const title = document.createElement('strong');
	title.textContent = 'Restrict editing';
	const closeButton = document.createElement('button');
	closeButton.type = 'button';
	closeButton.textContent = '×';
	closeButton.setAttribute('aria-label', 'Close');
	const body = document.createElement('div');
	body.className = 'dve-restrict-body';
	const style = document.createElement('style');
	style.textContent = STYLE;
	head.append(title, closeButton);
	const close = () => {
		element.hidden = true;
		// Closed, the pane holds no fields, so their names never collide with the document.
		element.replaceChildren();
		focusView(host.view());
	};
	const enforced = () => {
		const protection = host.protection();
		return Boolean(protection?.enforced && protection.edit !== 'none');
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
	const render = () => {
		say('');
		const protection = host.protection();
		const current = enforced();
		body.replaceChildren(...(current ? enforcedView : settingsView));
		for (const input of highlights) input.checked = host.highlight();
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
		localizeElement(element, locale);
		title.textContent = translate(locale, 'Restrict editing');
		closeButton.setAttribute('aria-label', translate(locale, 'Close'));
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
		render();
	});
	stop.addEventListener('click', () => {
		const protection = host.protection();
		const verdict = verifyProtectionPassword(protection, secret.value);
		if (verdict === undefined)
			return say('This protection uses a password that cannot be checked here');
		if (!verdict) return say('Incorrect password');
		// Stopping keeps what was chosen, as Word does, but no longer enforces it.
		host.apply(
			protection && {
				edit: protection.edit,
				enforced: false,
				...(protection.formatting && { formatting: true }),
			},
		);
		render();
	});
	closeButton.addEventListener('click', close);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
	});
	return {
		element,
		get isOpen() {
			return !element.hidden;
		},
		open() {
			element.replaceChildren(style, head, body);
			render();
			element.hidden = false;
			element.querySelector<HTMLElement>('input, select, button')?.focus();
		},
		close,
		setLocale(value) {
			locale = value;
			if (!element.hidden) render();
		},
	};
}
