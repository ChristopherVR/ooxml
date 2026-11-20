import type { SignatureIcon, SignatureStatusView, SignatureTone } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';
import { createEl } from '../render';
import { appendInfoDoneButton, openFileInfoDialogShell } from './file-info-dialog-shell';
import type { OverlayOwner } from './overlay-owner';

const TONE_CLASS: Readonly<Record<SignatureTone, string>> = {
	success: 'is-signed',
	caution: 'is-warning',
	danger: 'is-danger',
	neutral: '',
};

const ICON_GLYPH: Readonly<Record<SignatureIcon, string>> = {
	'shield-check': '✓',
	'shield-alert': '!',
	'shield-x': '×',
	'shield-question': '?',
};

function notice(doc: Document, className: string, glyph: string, ...lines: string[]): HTMLElement {
	const element = createEl(doc, 'div', `pptxv-info-notice ${className}`.trim());
	const icon = createEl(doc, 'b');
	icon.textContent = glyph;
	const copy = createEl(doc, 'div');
	for (const [index, line] of lines.entries()) {
		const paragraph = createEl(doc, 'p', index === 0 ? '' : 'pptxv-info-description');
		paragraph.textContent = line;
		copy.appendChild(paragraph);
	}
	element.append(icon, copy);
	return element;
}

/**
 * File > Info > Digital Signatures: the deck's verified signature state
 * (`signatureStatusView`), one row per signature, and the note that saving an
 * edit removes them. A deck reads as signed only when its signatures verify.
 */
export function openDigitalSignaturesDialog(
	doc: Document,
	t: Translator,
	view: SignatureStatusView,
	/** The viewer: the dialog closes with it. */
	owner?: OverlayOwner,
): HTMLElement {
	const shell = openFileInfoDialogShell(doc, t, t('pptx.digitalSignatures.title'), owner);
	shell.overlay.dataset.pptxSignatureState = view.state;
	if (!view.present) {
		shell.body.appendChild(notice(doc, '', 'i', t('pptx.digitalSignatures.noSignatures')));
		appendInfoDoneButton(doc, t, shell);
		return shell.overlay;
	}
	const status = notice(
		doc,
		TONE_CLASS[view.tone],
		ICON_GLYPH[view.icon],
		view.label,
		view.description,
	);
	status.dataset.pptxSignatureState = view.state;
	shell.body.appendChild(status);
	const list = createEl(doc, 'div', 'pptxv-info-list');
	for (const row of view.rows) {
		const item = createEl(doc, 'div', 'pptxv-info-row');
		item.dataset.pptxSignatureRowState = row.state;
		const details = createEl(doc, 'div');
		const signer = createEl(doc, 'strong');
		signer.textContent = row.signer;
		details.appendChild(signer);
		const lines = [
			row.issuer ? `${t('pptx.digitalSignatures.issuer')}: ${row.issuer}` : '',
			row.signedAt ? `${view.signedAtCaption}: ${row.signedAt}` : '',
		];
		for (const line of lines.filter(Boolean)) {
			const paragraph = createEl(doc, 'p', 'pptxv-info-description');
			paragraph.textContent = line;
			details.appendChild(paragraph);
		}
		const state = createEl(doc, 'em', `pptxv-info-signature-state ${TONE_CLASS[row.tone]}`.trim());
		state.textContent = row.label;
		item.append(details, state);
		list.appendChild(item);
	}
	shell.body.appendChild(list);
	shell.body.appendChild(notice(doc, 'is-warning', 'i', view.removedOnSave));
	appendInfoDoneButton(doc, t, shell);
	return shell.overlay;
}
