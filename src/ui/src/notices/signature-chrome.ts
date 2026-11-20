import type { signature } from 'ooxml-core/opc';
import { defineSignatureStatus, type OfficeUiSignatureStatus } from './signature-status';

/**
 * The status bar badge and the load notice a product shows for an opened document's signatures.
 * `update` is called with every render: a new summary (another document) shows the notice again
 * after Dismiss; the same summary keeps it dismissed. Both elements stay hidden for an unsigned
 * document. View signatures and the badge call `open` (the product's Info page); while that page
 * is unavailable (`setActivatable(false)`, a hidden File tab) both still show the verification
 * state but open nothing.
 */
export interface SignatureChrome {
	/** `office-ui-signature-status` for the status bar's default slot. */
	readonly badge: OfficeUiSignatureStatus;
	/** The `notice` variant, placed under the ribbon. */
	readonly notice: OfficeUiSignatureStatus;
	update(summary: signature.SignatureSummary | undefined, locale: string): void;
	/** False renders the badge and notice non-interactive (no View signatures); true restores them. */
	setActivatable(activatable: boolean): void;
}

/** The Info page's list of signatures: one `panel` element, or nothing for an unsigned document. */
export function signaturePanel(
	doc: Document,
	summary: signature.SignatureSummary | undefined,
	locale: string,
): HTMLElement[] {
	if (!summary || summary.state === 'unsigned') return [];
	defineSignatureStatus();
	const panel = doc.createElement('office-ui-signature-status');
	panel.variant = 'panel';
	panel.locale = locale;
	panel.summary = summary;
	panel.dataset.testid = 'signature-panel';
	return [panel];
}

export function createSignatureChrome(doc: Document, open: () => void): SignatureChrome {
	defineSignatureStatus();
	const make = (variant: 'badge' | 'notice') => {
		const element = doc.createElement('office-ui-signature-status');
		element.variant = variant;
		element.dataset.testid = `signature-${variant}`;
		element.addEventListener('office-signature-activate', () => {
			if (!element.noActivate) open();
		});
		return element;
	};
	const badge = make('badge');
	const notice = make('notice');
	let shown: signature.SignatureSummary | undefined;
	let dismissed = false;
	notice.addEventListener('office-signature-dismiss', () => {
		dismissed = true;
		notice.summary = undefined;
	});
	return {
		badge,
		notice,
		update(summary, locale) {
			if (summary !== shown) dismissed = false;
			shown = summary;
			badge.locale = notice.locale = locale;
			badge.summary = summary;
			notice.summary = dismissed ? undefined : summary;
		},
		setActivatable(activatable) {
			badge.noActivate = notice.noActivate = !activatable;
		},
	};
}
