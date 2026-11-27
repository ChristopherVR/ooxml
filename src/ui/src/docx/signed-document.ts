import type { confirmDialog } from '../dialog/message-dialogs';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';

const bars = new WeakMap<EditorCore, HTMLElement>();

/**
 * The confirmation dialog, passed in by the editor that already loaded it, so this chunk does not
 * import the message dialogs a second way (which would split them out of the shared chunk).
 */
type Confirm = typeof confirmDialog;

/**
 * Word's bar over a signed document: it opens read-only, and Edit Anyway warns that editing
 * removes the signatures before it lets the user edit. Shown while the signature lock holds.
 */
export function syncSignedDocumentBar(core: EditorCore, confirm: Confirm): void {
	const t = (key: string) => translate(core.locale, key as LocalizationKey);
	let bar = bars.get(core);
	if (!bar) {
		const anchor = core.shell.chrome?.signatures.notice ?? core.shell.toolbar;
		if (!anchor || !core.signatureLocked) return;
		bar = document.createElement('div');
		bar.className = 'dve-signed-bar';
		bar.setAttribute('role', 'status');
		bar.style.cssText = 'display:flex;gap:.75em;align-items:center;padding:.25em .75em';
		const text = document.createElement('span');
		const button = document.createElement('button');
		button.type = 'button';
		button.addEventListener('click', () => void editAnyway(core, confirm));
		bar.append(text, button);
		anchor.after(bar);
		bars.set(core, bar);
	}
	bar.hidden = !core.signatureLocked;
	bar.firstElementChild!.textContent = t(
		'This document has been signed and is read-only so that the signatures stay valid.',
	);
	bar.lastElementChild!.textContent = t('Edit Anyway');
}

/** Edit Anyway: after the user confirms, the signatures are removed and editing is allowed. */
export async function editAnyway(core: EditorCore, confirm: Confirm): Promise<void> {
	const root = core.element.shadowRoot ?? core.element;
	const confirmed = await confirm({
		title: translate(core.locale, 'Edit Anyway' as LocalizationKey),
		message: translate(
			core.locale,
			'Editing will remove the signatures from this document. Do you want to edit it?' as LocalizationKey,
		),
		confirmLabel: translate(core.locale, 'Edit Anyway' as LocalizationKey),
		host: root,
	});
	if (!confirmed || !core.signatureLocked) return;
	core.signatureLocked = false;
	// Saving a model without the signatures writes the package without them.
	const { signatures: _removed, ...rest } = core.model;
	core.model = rest;
	core.notifyChange();
	core.reportReadOnly();
	syncSignedDocumentBar(core, confirm);
}
