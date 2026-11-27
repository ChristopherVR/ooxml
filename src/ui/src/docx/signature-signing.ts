import type { SignatureLineSpec } from 'ooxml-core/docx';
import type { EditorCore } from './editor-core';
import { emit } from './events';
import { translate, type LocalizationKey } from './localization';

const VERIFIED = new Set(['verified', 'verified-untrusted']);

const parseSpec = (json: string): SignatureLineSpec | undefined => {
	try {
		const spec = JSON.parse(json) as SignatureLineSpec;
		return typeof spec?.id === 'string' ? spec : undefined;
	} catch {
		return undefined;
	}
};

/** The verified signature that signs the line `id`, if any. */
export function signingOf(core: EditorCore, id: string) {
	return core.model.signatures?.signatures.find(
		(entry) => VERIFIED.has(entry.state) && entry.signatureLineId === id,
	);
}

const commonName = (subject: string | undefined) =>
	(/(?:^|,\s*)CN=([^,]+)/u.exec(subject ?? '')?.[1] ?? subject ?? '').trim();

/**
 * The signer's name or picture over a signed line, with "Digitally signed by" and the date, as
 * Word draws a signed signature line. Undefined for a line no verified signature signs.
 */
export function signedLineOverlay(core: EditorCore, json: string): HTMLElement | undefined {
	const spec = parseSpec(json);
	const signing = spec && signingOf(core, spec.id);
	if (!spec || !signing) return undefined;
	const overlay = document.createElement('span');
	overlay.className = 'dve-signed-line';
	overlay.setAttribute('aria-hidden', 'true');
	overlay.style.cssText =
		'position:absolute;inset:0 0 30% 0;display:flex;flex-direction:column;justify-content:flex-end;' +
		'align-items:flex-start;padding:0 6%;background:#fff;color:#000;line-height:1.2;pointer-events:none';
	if (signing.signatureImage?.startsWith('iVBOR')) {
		const image = document.createElement('img');
		image.src = `data:image/png;base64,${signing.signatureImage}`;
		image.alt = '';
		image.style.cssText = 'max-height:70%;max-width:100%;object-fit:contain';
		overlay.append(image);
	} else {
		const name = document.createElement('span');
		name.textContent = signing.signatureText || commonName(signing.signer);
		name.style.cssText = "font:italic 22px 'Segoe Script','Brush Script MT',cursive";
		overlay.append(name);
	}
	const by = document.createElement('small');
	const date =
		spec.showSignDate && signing.signingTime
			? ` ${new Date(signing.signingTime).toLocaleDateString(core.locale)}`
			: '';
	by.textContent = `${translate(core.locale, 'Digitally signed by' as LocalizationKey)} ${commonName(signing.signer)}${date}`;
	by.style.cssText = 'font:10px sans-serif;opacity:.8';
	overlay.append(by);
	return overlay;
}

/**
 * Double-click on a signature line (or Insert > Signature Line > Sign): the Sign dialog, which
 * needs the host's `documentSigner`. A signed line shows who signed it instead.
 */
export async function openSignDialog(core: EditorCore, json?: string): Promise<void> {
	await import('./locales/insert-parts-translations');
	const spec = json ? parseSpec(json) : undefined;
	const say = (message: string) =>
		emit(core.element, 'document-warning', translate(core.locale, message as LocalizationKey));
	if (spec && signingOf(core, spec.id)) {
		say('This signature line has been signed.');
		return;
	}
	if (!core.documentSigner) {
		say('Signing needs a certificate, which the host of this editor has not provided.');
		return;
	}
	if (core.collab.active) {
		say('Stop sharing the document before signing it.');
		return;
	}
	const { createSignDialog } = await import('./signature-sign-dialog');
	createSignDialog(core, spec);
}
