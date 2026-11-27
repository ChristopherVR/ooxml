import { signatureLinesOf, type SignatureLineSpec } from 'ooxml-core/docx';
// The signer itself loads on the first signing (`signature.signPackage`), so this module stays small.
import { signature } from 'ooxml-core/opc';
import { dialogButton, labelled, textInput } from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';
import { signingOf } from './signature-signing';

const MAX_IMAGE = { width: 600, height: 200 };

/** A chosen picture as a PNG no larger than {@link MAX_IMAGE}, Base64. */
async function pictureAsPng(file: File): Promise<string | undefined> {
	if (!/^image\/(png|jpeg|gif|bmp|webp)$/u.test(file.type) || file.size > 8 * 1024 * 1024)
		return undefined;
	const bitmap = await createImageBitmap(file);
	const scale = Math.min(1, MAX_IMAGE.width / bitmap.width, MAX_IMAGE.height / bitmap.height);
	const canvas = document.createElement('canvas');
	canvas.width = Math.max(1, Math.round(bitmap.width * scale));
	canvas.height = Math.max(1, Math.round(bitmap.height * scale));
	canvas.getContext('2d')?.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
	bitmap.close();
	return canvas.toDataURL('image/png').replace(/^data:image\/png;base64,/u, '');
}

/**
 * Word's Sign dialog: the signer types a name or chooses a picture, adds comments where the line
 * allows them, and signs with the host's certificate. The document is saved, signed as a whole
 * (an XML digital signature naming this line) and reopened signed and read-only, as Word does.
 */
export function createSignDialog(core: EditorCore, requested?: SignatureLineSpec): void {
	const root = core.element.shadowRoot;
	const signer = core.documentSigner;
	if (!root || !signer) return;
	const t = (key: string) => translate(core.locale, key as LocalizationKey);
	const line =
		requested ??
		signatureLinesOf(core.model.blocks).find((candidate) => !signingOf(core, candidate.id));
	const name = textInput();
	name.maxLength = 255;
	const comments = textInput();
	comments.maxLength = 255;
	let image: string | undefined;
	const picker = document.createElement('input');
	picker.type = 'file';
	picker.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp';
	picker.hidden = true;
	const choose = dialogButton('Select Image…');
	const chosen = document.createElement('span');
	choose.addEventListener('click', () => picker.click());
	picker.addEventListener('change', async () => {
		const file = picker.files?.[0];
		image = file ? await pictureAsPng(file).catch(() => undefined) : undefined;
		chosen.textContent = image ? (file?.name ?? '') : '';
		validate();
	});
	const status = document.createElement('p');
	status.setAttribute('role', 'status');
	const sign = dialogButton('Sign', true);
	const cancel = dialogButton('Cancel');
	const validate = () => {
		sign.disabled = Boolean(line) && !name.value.trim() && !image;
	};
	name.addEventListener('input', validate);
	const element = createDialogShell('Sign', 'dve-format-dialog dve-sign-dialog');
	const close = () => {
		setDialogOpen(element, false);
		element.remove();
		core.targetView()?.focus();
	};
	const paragraph = (text: string) => {
		const p = document.createElement('p');
		p.textContent = text;
		return p;
	};
	const content: HTMLElement[] = [];
	if (line?.instructions) content.push(paragraph(line.instructions));
	if (line) {
		content.push(
			paragraph('Type your name below, or select an image to use as a graphical signature:'),
			labelled('Signer name', name),
			choose,
			chosen,
			picker,
		);
		if (line.suggestedSigner)
			content.push(paragraph(`${t('Suggested signer')}: ${line.suggestedSigner}`));
	} else content.push(paragraph('Sign this document with a digital signature.'));
	if (!line || line.allowComments)
		content.push(labelled(line ? 'Comments' : 'Purpose for signing this document', comments));
	content.push(
		paragraph(`${t('Signing as')}: ${signer.label ?? t('the certificate the host provides')}`),
		paragraph('The document is signed as it is now. Editing it afterwards removes the signature.'),
		status,
		dialogActions(cancel, sign),
	);
	cancel.addEventListener('click', close);
	onDialogDismiss(element, close);
	sign.addEventListener('click', async () => {
		sign.disabled = true;
		status.textContent = t('Signing…');
		try {
			const bytes = await (
				core.element as unknown as { saveBytes(): Promise<Uint8Array> }
			).saveBytes();
			const signed = await signature.signPackage(bytes, signer, {
				...(line ? { signatureLineId: line.id } : {}),
				...(name.value.trim() ? { signatureText: name.value.trim() } : {}),
				...(image ? { signatureImage: image } : {}),
				...(comments.value.trim() ? { comments: comments.value.trim() } : {}),
				...(line?.suggestedSignerTitle ? { signerRole: line.suggestedSignerTitle } : {}),
			});
			close();
			await (core.element as unknown as { load(bytes: Uint8Array): Promise<void> }).load(signed);
			// The signed file is what the host should keep: it is unsaved until the host saves it.
			core.dirtyState.set(true);
			core.shell.chrome?.setSaveState('dirty');
		} catch {
			status.textContent = t('The document could not be signed.');
			sign.disabled = false;
		}
	});
	element.replaceChildren(...content);
	localizeDialog(element, core.locale);
	root.append(element);
	validate();
	setDialogOpen(element, true);
	(line ? name : comments).focus();
}
