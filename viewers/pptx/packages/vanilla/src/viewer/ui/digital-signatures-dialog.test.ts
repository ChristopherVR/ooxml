import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n';
import { openDigitalSignaturesDialog } from './digital-signatures-dialog';

afterEach(() => document.body.replaceChildren());

function summaryOf(state: SignatureState): SignatureSummary {
	return {
		state,
		signatures: [
			{
				path: '_xmlsignatures/sig1.xml',
				state,
				signer: 'CN=Ada Lovelace, O=Analytical',
				issuer: 'CN=Test Root',
				signingTime: '2026-01-02T03:04:05Z',
			},
		],
	};
}

function open(summary: SignatureSummary | undefined): HTMLElement {
	return openDigitalSignaturesDialog(
		document,
		createTranslator(),
		signatureStatusView(summary, 'en'),
	);
}

/** The status notice's heading line: what the user reads first. */
function headline(overlay: HTMLElement): string {
	return overlay.querySelector('.pptxv-info-notice p')!.textContent ?? '';
}

describe('digital signatures dialog', () => {
	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
		['verified-untrusted', 'Signed, signer not verified'],
		['verified', 'Signed'],
	] as const)('shows a %s deck as "%s"', (state, label) => {
		const overlay = open(summaryOf(state));
		expect(overlay.dataset.pptxSignatureState).toBe(state);
		expect(headline(overlay)).toBe(label);
		const row = overlay.querySelector<HTMLElement>('[data-pptx-signature-row-state]')!;
		expect(row.dataset.pptxSignatureRowState).toBe(state);
		expect(row.querySelector('strong')!.textContent).toBe('Ada Lovelace');
		expect(row.querySelector('.pptxv-info-signature-state')!.textContent).toBe(label);
		expect(row.textContent).toContain('Issuer: CN=Test Root');
		expect(row.textContent).toContain('Signed on: ');
	});

	it('never labels an unverified signature "Signed"', () => {
		for (const state of ['invalid', 'unsupported'] as const) {
			const overlay = open(summaryOf(state));
			expect(headline(overlay)).not.toMatch(/^Signed/);
			expect(overlay.textContent).not.toContain('This presentation is digitally signed.');
		}
	});

	it('notes that saving removes the signatures', () => {
		const overlay = open(summaryOf('verified'));
		expect(overlay.textContent).toContain(
			'Saving changes removes the signatures, because any change invalidates them.',
		);
	});

	it('says an unsigned deck has no signatures', () => {
		const overlay = open(undefined);
		expect(overlay.dataset.pptxSignatureState).toBe('unsigned');
		expect(headline(overlay)).toBe('This presentation has no digital signatures.');
		expect(overlay.querySelector('[data-pptx-signature-row-state]')).toBeNull();
	});
});
