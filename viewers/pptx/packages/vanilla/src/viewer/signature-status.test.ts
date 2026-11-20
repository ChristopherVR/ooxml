// @vitest-own-worker: mounts full viewers; see region-customization.test-support.ts.
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

import type { SignatureSummary } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { destroyMounted, internals, mount, storeOf } from './region-customization.test-support';

afterEach(destroyMounted);

const invalid: SignatureSummary = {
	state: 'invalid',
	signatures: [{ path: '_xmlsignatures/sig1.xml', state: 'invalid', signer: 'CN=Ada' }],
};

function summary(state: SignatureSummary['state']): SignatureSummary {
	return { state, signatures: [{ path: '_xmlsignatures/sig1.xml', state, signer: 'CN=Ada' }] };
}

function signaturesDialog(): HTMLElement {
	return document.querySelector<HTMLElement>('.pptxv-info-overlay[data-pptx-signature-state]')!;
}

function badge(): HTMLButtonElement | null {
	return document.querySelector<HTMLButtonElement>('.pptxv-signature-badge');
}

describe('vanilla signature state', () => {
	it('opens File > Info > Signatures with the loaded deck state in the UI locale', () => {
		const { viewer } = mount({ editable: true });
		storeOf(viewer).set({ signatures: invalid });
		internals(viewer).openDigitalSignatures();
		expect(signaturesDialog().dataset.pptxSignatureState).toBe('invalid');
		expect(signaturesDialog().textContent).toContain('Invalid signature');
		signaturesDialog().remove();
		viewer.setLocale('de');
		internals(viewer).openDigitalSignatures();
		expect(signaturesDialog().textContent).not.toContain('Invalid signature');
	});

	it('shows the badge under the Ribbon only when the signatures verify as signed', () => {
		const { viewer } = mount({ editable: true });
		expect(badge()).toBeNull();
		const cases: [SignatureSummary['state'], string][] = [
			['invalid', 'Invalid signature'],
			['unsupported', 'Signature not verified'],
			['verified-untrusted', 'Signed, signer not verified'],
			['verified', 'Signed'],
		];
		for (const [state, label] of cases) {
			storeOf(viewer).set({ signatures: summary(state) });
			expect(badge()?.dataset.pptxSignatureState).toBe(state);
			expect(badge()?.textContent).toBe(label);
		}
		storeOf(viewer).set({ signatures: undefined });
		expect(badge()).toBeNull();
	});

	it('opens Digital Signatures from the badge and words it in the new locale', () => {
		const { viewer } = mount({ editable: true });
		storeOf(viewer).set({ signatures: invalid });
		badge()!.click();
		expect(signaturesDialog().dataset.pptxSignatureState).toBe('invalid');
		signaturesDialog().remove();
		viewer.setLocale('de');
		expect(badge()?.dataset.pptxSignatureState).toBe('invalid');
		expect(badge()?.textContent).not.toBe('Invalid signature');
	});

	it('keeps the badge but opens no dialog when the host hides the signatures dialog', () => {
		const { viewer } = mount({ editable: true, customization: { hiddenDialogs: ['signatures'] } });
		storeOf(viewer).set({ signatures: invalid });
		badge()!.click();
		expect(signaturesDialog()).toBeNull();
	});

	it('warns on the first edit of a deck whose signature does not verify', () => {
		const { viewer } = mount({ editable: true });
		storeOf(viewer).set({ signatures: invalid });
		storeOf(viewer).set({ dirty: true });
		expect(document.body.textContent).toContain('Remove signatures?');
	});

	it('opens a signed deck with a dismissible notice of its state, worded in the UI locale', async () => {
		const deck = (name: string) =>
			readFileSync(
				join(import.meta.dirname, '../../../../../../e2e/pptx/fixtures/signature-states', name),
			);
		const { viewer } = mount({ editable: true, showToolbar: false });
		const notice = () =>
			storeOf(viewer)
				.get()
				.compatToasts.find((toast) => toast.code === 'SIGNATURE_STATE');
		await viewer.loadFile(deck('invalid.pptx'));
		expect(notice()?.severity).toBe('warning');
		expect(notice()?.text).toContain('is not valid');
		expect(notice()?.text).toContain('Saving changes removes the signatures');
		viewer.setLocale('de');
		expect(notice()?.text).not.toContain('Saving changes');
		document
			.querySelector('pptx-ui-compat-toasts')
			?.shadowRoot?.querySelector<HTMLButtonElement>(
				'[data-code="SIGNATURE_STATE"] [data-testid="pptx-compat-toast-dismiss"]',
			)
			?.click();
		expect(notice()).toBeUndefined();
		await viewer.loadFile(deck('unsigned.pptx'));
		expect(notice()).toBeUndefined();
		await viewer.loadFile(deck('verified-untrusted.pptx'));
		expect(notice()?.severity).toBe('info');
	});
});
