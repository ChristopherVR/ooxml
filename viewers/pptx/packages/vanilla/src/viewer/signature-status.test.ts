// @vitest-own-worker: mounts full viewers; see region-customization.test-support.ts.
import type { SignatureSummary } from 'pptx-viewer-core';
import { afterEach, describe, expect, it } from 'vitest';

import { destroyMounted, internals, mount, storeOf } from './region-customization.test-support';

afterEach(destroyMounted);

const invalid: SignatureSummary = {
	state: 'invalid',
	signatures: [{ path: '_xmlsignatures/sig1.xml', state: 'invalid', signer: 'CN=Ada' }],
};

function signaturesDialog(): HTMLElement {
	return document.querySelector<HTMLElement>('.pptxv-info-overlay[data-pptx-signature-state]')!;
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

	it('warns on the first edit of a deck whose signature does not verify', () => {
		const { viewer } = mount({ editable: true });
		storeOf(viewer).set({ signatures: invalid });
		storeOf(viewer).set({ dirty: true });
		expect(document.body.textContent).toContain('Remove signatures?');
	});
});
