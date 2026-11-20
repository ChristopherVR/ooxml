import { html, nothing } from 'lit';
import { signatureStateText, signatureStringsFor } from 'ooxml-core/i18n';
import { signature as opcSignature } from 'ooxml-core/opc';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { getIcon, registerIcon } from '../icons';
import { definer } from '../registry';
import css from './signature-status.css?raw';

type SignatureSummary = opcSignature.SignatureSummary;
type SignatureEntry = opcSignature.SignatureEntry;

/** How the element draws a summary: a status bar entry, a load notice or an info panel. */
export type OfficeSignatureStatusVariant = 'badge' | 'notice' | 'panel';

// A certificate rosette, the mark Office puts on a signed document's status bar.
const SIGNATURE_ICON = 'M6 7.5a4 4 0 1 0 8 0 4 4 0 1 0-8 0M7.6 10.8 6.5 17l3.5-2 3.5 2-1.1-6.2';

const formatTime = (value: string | undefined, locale: string): string | undefined => {
	const time = value ? Date.parse(value) : Number.NaN;
	if (!Number.isFinite(time)) return undefined;
	try {
		return new Intl.DateTimeFormat(locale, { dateStyle: 'medium', timeStyle: 'short' }).format(
			time,
		);
	} catch {
		return new Date(time).toISOString();
	}
};

/**
 * `<office-ui-signature-status>`: a document's digital signature state, as Word, Excel and Visio
 * show it. `summary` is the `SignatureSummary` a load returned (`signature` of `ooxml-core/opc`); `locale` picks
 * the shared wording of `ooxml-core/i18n` (`signatureStateText`), so the products never describe a
 * state differently. Only `verified` reads "Signed": a signature that does not verify is shown as
 * invalid or not verified. Hidden while `summary` is absent or `unsigned`. The state is mirrored to
 * the host's `data-state`.
 *
 * `variant`: `badge` (default) is a status bar button that emits `office-signature-activate`;
 * `notice` is a load message bar with View signatures (`office-signature-activate`) and Dismiss
 * (`office-signature-dismiss`); `panel` lists every signature with its signer and signing time for
 * an info page. Event names are static fields a product subclass may override.
 *
 * `no-activate` keeps the state on screen but drops what it would open: the badge is plain text
 * and the notice has no View signatures (Dismiss stays). A product sets it while the page the
 * badge opens is unavailable, for example when a host hid the File tab.
 */
export class OfficeUiSignatureStatus extends OfficeElement {
	static activateEvent = 'office-signature-activate';
	static dismissEvent = 'office-signature-dismiss';
	static override styles = controlStyles(css);
	static override properties = {
		summary: { attribute: false },
		locale: { type: String },
		variant: { type: String, reflect: true },
		noActivate: { attribute: 'no-activate', ...flag },
	};
	declare summary: SignatureSummary | undefined;
	declare locale: string;
	declare variant: OfficeSignatureStatusVariant;
	declare noActivate: boolean;

	constructor() {
		super();
		if (!getIcon('signature')) registerIcon('signature', SIGNATURE_ICON);
		this.summary = undefined;
		this.locale = 'en';
		this.variant = 'badge';
		this.noActivate = false;
	}

	protected override willUpdate(): void {
		const state = this.summary?.state ?? 'unsigned';
		this.setAttribute('data-state', state);
		this.toggleAttribute('hidden', state === 'unsigned');
	}

	private send(name: 'activateEvent' | 'dismissEvent'): void {
		this.fire((this.constructor as typeof OfficeUiSignatureStatus)[name], {
			state: this.summary?.state ?? 'unsigned',
		});
	}

	private entry(signature: SignatureEntry) {
		const words = signatureStringsFor(this.locale);
		const signer = opcSignature.signerDisplayName(signature.signer);
		const time = formatTime(signature.signingTime, this.locale);
		const state = signatureStateText(signature.state, this.locale);
		return html`<li class="entry" data-state=${signature.state}
			>${glyph('signature', 'icon')}<div
				><p class="entry-state">${state?.label ?? ''}</p>${
					signer
						? html`<p
								><span class="caption">${words.signer}</span>
								<span class="signer">${signer}</span></p
							>`
						: nothing
				}${
					time
						? html`<p><span class="caption">${words.signedAt}</span> <span>${time}</span></p>`
						: nothing
				}</div
			></li
		>`;
	}

	protected override render() {
		const summary = this.summary;
		const text = summary ? signatureStateText(summary.state, this.locale) : undefined;
		if (!summary || !text) return nothing;
		const words = signatureStringsFor(this.locale);
		if (this.variant === 'notice')
			return html`<div class="notice" role="status" part="notice"
				>${glyph('signature', 'icon')}<p class="text"
					><strong class="label">${text.label}</strong>
					<span class="description">${text.description} ${words.removedOnSave}</span></p
				>${
					this.noActivate
						? nothing
						: html`<button type="button" class="view" @click=${() => this.send('activateEvent')}
								>${words.view}</button
							>`
				}<button
					type="button"
					class="dismiss"
					title=${words.dismiss}
					aria-label=${words.dismiss}
					@click=${() => this.send('dismissEvent')}
					>${glyph('close', 'icon')}</button
				></div
			>`;
		if (this.variant === 'panel')
			return html`<section class="panel" part="panel" aria-label=${words.heading}
				><h3 class="heading">${words.heading}</h3
				><p class="state"
					>${glyph('signature', 'icon')}<strong class="label">${text.label}</strong></p
				><p class="description">${text.description}</p>${
					summary.signatures.length
						? html`<ul class="entries"
								>${summary.signatures.map((signature) => this.entry(signature))}</ul
							>`
						: nothing
				}<p class="removed">${words.removedOnSave}</p></section
			>`;
		if (this.noActivate)
			return html`<span class="badge" part="badge" title=${text.description}
				>${glyph('signature', 'icon')}<span class="hidden-text">${words.heading}: </span
				><span class="label">${text.label}</span></span
			>`;
		return html`<button
			type="button"
			class="badge"
			part="badge"
			title=${text.description}
			aria-label=${`${words.heading}: ${text.label}`}
			@click=${() => this.send('activateEvent')}
			>${glyph('signature', 'icon')}<span class="label">${text.label}</span></button
		>`;
	}
}

export const defineSignatureStatus = definer(
	'office-ui-signature-status',
	() => OfficeUiSignatureStatus,
);

declare global {
	interface HTMLElementTagNameMap {
		'office-ui-signature-status': OfficeUiSignatureStatus;
	}
}
