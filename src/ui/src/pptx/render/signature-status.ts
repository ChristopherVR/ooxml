import { signatureStringsFor } from 'ooxml-core/i18n';
import {
	isVerifiedSignatureState,
	signerDisplayName,
	type SignatureEntry,
	type SignatureState,
	type SignatureSummary,
} from 'ooxml-core/pptx';

import type { CompatibilityWarningToast } from './compatibility-warning-toasts';

/**
 * How every PowerPoint surface (badge, status bar entry, info panel, signatures panel, load
 * notice) presents a deck's digital signatures. A deck reads as signed only when its signatures
 * verify; a signature part alone never does. Bindings map this descriptor onto their templates
 * and keep no status logic of their own.
 */

/** The colour family of a state: a binding maps it onto its own palette. */
export type SignatureTone = 'success' | 'caution' | 'danger' | 'neutral';

/** The icon of a state: a binding maps it onto its own icon set. */
export type SignatureIcon = 'shield-check' | 'shield-alert' | 'shield-x' | 'shield-question';

const LOOK: Readonly<Record<SignatureState, { tone: SignatureTone; icon: SignatureIcon }>> = {
	unsigned: { tone: 'neutral', icon: 'shield-question' },
	verified: { tone: 'success', icon: 'shield-check' },
	'verified-untrusted': { tone: 'caution', icon: 'shield-alert' },
	invalid: { tone: 'danger', icon: 'shield-x' },
	unsupported: { tone: 'neutral', icon: 'shield-question' },
};

/** One signature as a signatures panel lists it. */
export interface SignatureStatusRow {
	/** The signature part, unique in the package: use it as the list key. */
	path: string;
	state: SignatureState;
	/** The state's short label, for the row's badge. */
	label: string;
	tone: SignatureTone;
	icon: SignatureIcon;
	/** The signer's common name, else the signature part's path. */
	signer: string;
	/** The full certificate subject, when it could be read. */
	subject?: string;
	issuer?: string;
	/** The signing time the signature claims, formatted for the locale. */
	signedAt?: string;
	/** ISO dates of the certificate's validity. */
	validFrom?: string;
	validTo?: string;
}

/** The whole deck's signature presentation in one locale. */
export interface SignatureStatusView {
	state: SignatureState;
	/** Signature parts are present: show the badge and the list (saving removes them). */
	present: boolean;
	/** The deck is shown as signed: every signature verified. */
	signed: boolean;
	/** A few words for a badge or status bar entry; empty for an unsigned deck. */
	label: string;
	/** A sentence for a tooltip, info panel or load notice; empty for an unsigned deck. */
	description: string;
	tone: SignatureTone;
	icon: SignatureIcon;
	count: number;
	rows: SignatureStatusRow[];
	/** The panel heading and captions, and the note that saving an edit removes the signatures. */
	heading: string;
	signerCaption: string;
	signedAtCaption: string;
	removedOnSave: string;
}

function formatInstant(iso: string, locale: string | undefined): string {
	const date = new Date(iso);
	if (Number.isNaN(date.getTime())) return iso;
	try {
		return date.toLocaleString(locale);
	} catch {
		return date.toLocaleString();
	}
}

function rowOf(entry: SignatureEntry, locale: string | undefined): SignatureStatusRow {
	const strings = signatureStringsFor(locale);
	return {
		path: entry.path,
		state: entry.state,
		label: entry.state === 'unsigned' ? '' : strings.states[entry.state].label,
		...LOOK[entry.state],
		signer: signerDisplayName(entry.signer) ?? entry.path,
		...(entry.signer ? { subject: entry.signer } : {}),
		...(entry.issuer ? { issuer: entry.issuer } : {}),
		...(entry.signingTime ? { signedAt: formatInstant(entry.signingTime, locale) } : {}),
		...(entry.validFrom ? { validFrom: entry.validFrom } : {}),
		...(entry.validTo ? { validTo: entry.validTo } : {}),
	};
}

/** The signature presentation of a deck (`PptxData.signatures`, absent when unsigned). */
export function signatureStatusView(
	summary: SignatureSummary | undefined,
	locale: string | undefined,
): SignatureStatusView {
	const strings = signatureStringsFor(locale);
	const state = summary?.state ?? 'unsigned';
	const text = state === 'unsigned' ? undefined : strings.states[state];
	const rows = (summary?.signatures ?? []).map((entry) => rowOf(entry, locale));
	return {
		state,
		present: state !== 'unsigned',
		signed: isVerifiedSignatureState(state),
		label: text?.label ?? '',
		description: text?.description ?? '',
		...LOOK[state],
		count: rows.length,
		rows,
		heading: strings.heading,
		signerCaption: strings.signer,
		signedAtCaption: strings.signedAt,
		removedOnSave: strings.removedOnSave,
	};
}

/** The id (and code) of the signature load notice in a binding's notice stack. */
export const SIGNATURE_NOTICE_ID = 'SIGNATURE_STATE';

/**
 * The notice a binding shows when a signed deck loads: the state's description and the note that
 * saving an edit removes the signatures, already translated. Undefined for an unsigned deck. A
 * binding adds it to its notice stack and hides it once dismissed, until the next deck loads.
 */
export function signatureLoadNotice(
	view: SignatureStatusView,
): CompatibilityWarningToast | undefined {
	if (!view.present) return undefined;
	return {
		id: SIGNATURE_NOTICE_ID,
		code: SIGNATURE_NOTICE_ID,
		severity: view.signed ? 'info' : 'warning',
		messageKey: '',
		text: `${view.description} ${view.removedOnSave}`,
	};
}
