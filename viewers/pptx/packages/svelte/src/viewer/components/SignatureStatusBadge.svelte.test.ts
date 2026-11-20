import { registerPptxWebControls, signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import DigitalSignaturesDialog from './ribbon/file/DigitalSignaturesDialog.svelte';
import SignatureStatusBadge from './SignatureStatusBadge.svelte';

registerPptxWebControls();

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	document.body.replaceChildren();
});

function open(component: never, props: Record<string, unknown>): HTMLElement {
	const target = document.createElement('div');
	document.body.append(target);
	const instance = mount(component, { target, props });
	flushSync();
	cleanup = () => unmount(instance);
	return target;
}

const summary = (state: SignatureState): SignatureSummary => ({
	state,
	signatures: [
		{
			path: '_xmlsignatures/sig1.xml',
			state,
			signer: 'CN=Ada Lovelace, O=Example',
			issuer: 'CN=Example CA',
		},
	],
});

const badgeOf = (state: SignatureState, onclick?: () => void): HTMLButtonElement | null =>
	open(SignatureStatusBadge as never, {
		view: signatureStatusView(summary(state), 'en'),
		onclick,
	}).querySelector('button');

describe('signature status badge', () => {
	it.each([
		['verified', 'Signed'],
		['verified-untrusted', 'Signed, signer not verified'],
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
	] as const)('labels a %s deck "%s"', (state, label) => {
		const badge = badgeOf(state);
		expect(badge?.textContent?.trim()).toBe(label);
		expect(badge?.dataset.pptxSignatureState).toBe(state);
	});

	it('never reads "Signed" for a signature that did not verify', () => {
		for (const state of ['invalid', 'unsupported'] as const) {
			expect(badgeOf(state)?.textContent?.trim()).not.toBe('Signed');
		}
	});

	it('renders nothing for an unsigned deck', () => {
		const target = open(SignatureStatusBadge as never, {
			view: signatureStatusView(undefined, 'en'),
		});
		expect(target.querySelector('[data-pptx-signature-state]')).toBeNull();
	});

	it('opens the signatures dialog on click', () => {
		const onclick = vi.fn();
		badgeOf('invalid', onclick)?.click();
		expect(onclick).toHaveBeenCalledOnce();
	});

	it('words the badge in the current locale', () => {
		const target = open(SignatureStatusBadge as never, {
			view: signatureStatusView(summary('verified'), 'fr'),
		});
		expect(target.querySelector('button')?.textContent?.trim()).toBe('Signé');
	});
});

describe('digital signatures dialog', () => {
	it('shows the verification state of the deck and of each signature', () => {
		const target = open(DigitalSignaturesDialog as never, {
			view: signatureStatusView(summary('invalid'), 'en'),
			onclose: () => undefined,
		});
		const panel = target.querySelector('[role="dialog"]');
		expect(panel?.getAttribute('data-pptx-signature-state')).toBe('invalid');
		expect(panel?.textContent).toContain('Invalid signature');
		expect(panel?.textContent).not.toMatch(/\bSigned\b/u);
		const row = target.querySelector('[data-pptx-signature-row-state]');
		expect(row?.getAttribute('data-pptx-signature-row-state')).toBe('invalid');
		expect(row?.textContent).toContain('Ada Lovelace');
		expect(row?.textContent).toContain('CN=Example CA');
	});

	it('reads "Signed, signer not verified" for a verified-untrusted deck', () => {
		const target = open(DigitalSignaturesDialog as never, {
			view: signatureStatusView(summary('verified-untrusted'), 'en'),
			onclose: () => undefined,
		});
		expect(target.querySelector('.notice strong')?.textContent).toBe('Signed, signer not verified');
	});

	it('says there are no signatures for an unsigned deck', () => {
		const target = open(DigitalSignaturesDialog as never, {
			view: signatureStatusView(undefined, 'en'),
			onclose: () => undefined,
		});
		expect(target.querySelector('[data-pptx-signature-row-state]')).toBeNull();
		expect(target.textContent).toContain('This presentation has no digital signatures.');
	});
});
