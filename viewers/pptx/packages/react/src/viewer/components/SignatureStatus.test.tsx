// @vitest-environment happy-dom
import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	useCompatibilityToastsState,
	useSignatureNotice,
} from '../hooks/useCompatibilityToastsState';
import { PptxI18nProvider } from '../i18n/provider';
import { CompatibilityToasts } from './CompatibilityToasts';
import { DigitalSignaturesDialog } from './DigitalSignaturesDialog';
import { SignatureStatusBadge } from './SignatureStatusBadge';

/**
 * A deck reads "Signed" only when its signatures verify: the badge and the signatures dialog
 * show the verified state of `PptxData.signatures`, never the mere presence of signature parts.
 */

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const deck = (state: SignatureState): SignatureSummary => ({
	state,
	signatures: [
		{
			path: '_xmlsignatures/sig1.xml',
			state,
			signer: 'CN=Ada Lovelace, O=Example',
			issuer: 'CN=Example CA',
			signingTime: '2026-01-02T03:04:05Z',
		},
	],
});

const badge = () => container.querySelector<HTMLButtonElement>('[data-pptx-signature-state]');

describe('signatureStatusBadge', () => {
	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
		['verified-untrusted', 'Signed, signer not verified'],
		['verified', 'Signed'],
	] as const)('labels a %s deck "%s"', (state, label) => {
		act(() => root.render(<SignatureStatusBadge signatures={deck(state)} onClick={() => {}} />));
		expect(badge()?.getAttribute('data-pptx-signature-state')).toBe(state);
		expect(badge()?.textContent).toBe(label);
		expect(badge()?.getAttribute('title')).not.toBe('');
	});

	it('renders nothing for an unsigned deck', () => {
		act(() => root.render(<SignatureStatusBadge signatures={undefined} onClick={() => {}} />));
		expect(container.innerHTML).toBe('');
	});

	it('opens the signatures dialog on click', () => {
		const onClick = vi.fn();
		act(() => root.render(<SignatureStatusBadge signatures={deck('invalid')} onClick={onClick} />));
		act(() => badge()?.click());
		expect(onClick).toHaveBeenCalledOnce();
	});

	it("speaks the viewer's language", () => {
		act(() =>
			root.render(
				<PptxI18nProvider translate={() => undefined} locale='fr'>
					<SignatureStatusBadge signatures={deck('invalid')} onClick={() => {}} />
				</PptxI18nProvider>,
			),
		);
		expect(badge()?.textContent).toBe('Signature non valide');
	});
});

describe('digitalSignaturesDialog', () => {
	const render = (signatures: SignatureSummary | undefined) =>
		act(() =>
			root.render(<DigitalSignaturesDialog isOpen onClose={() => {}} signatures={signatures} />),
		);

	it('never reads "Signed" for an invalid deck', () => {
		render(deck('invalid'));
		const panel = container.querySelector('[data-pptx-signature-state]');
		expect(panel?.getAttribute('data-pptx-signature-state')).toBe('invalid');
		const text = panel?.textContent ?? '';
		expect(text).toContain('Invalid signature');
		expect(text).not.toContain('This presentation is digitally signed.');
		const row = container.querySelector('[data-pptx-signature-row-state]');
		expect(row?.getAttribute('data-pptx-signature-row-state')).toBe('invalid');
		expect(row?.textContent).toContain('Ada Lovelace');
		expect(row?.textContent).toContain('CN=Example CA');
		expect(text).toContain('Saving changes removes the signatures');
	});

	it('reads "Signed, signer not verified" for a verified-untrusted deck', () => {
		render(deck('verified-untrusted'));
		expect(
			container.querySelector('[data-pptx-signature-row-state="verified-untrusted"]')?.textContent,
		).toContain('Signed, signer not verified');
	});

	it('shows the no-signatures notice for an unsigned deck', () => {
		render(undefined);
		const panel = container.querySelector('[data-pptx-signature-state]');
		expect(panel?.getAttribute('data-pptx-signature-state')).toBe('unsigned');
		expect(panel?.textContent).toContain('This presentation has no digital signatures.');
		expect(container.querySelector('[data-pptx-signature-row-state]')).toBeNull();
	});
});

describe('signature load notice', () => {
	function Stack({ signatures }: { signatures: SignatureSummary | undefined }) {
		const stack = useSignatureNotice(signatures, useCompatibilityToastsState());
		return (
			<CompatibilityToasts
				toasts={stack.toasts}
				onDismiss={stack.dismiss}
				onDismissAll={stack.dismissAll}
			/>
		);
	}
	const notice = () =>
		container
			.querySelector('pptx-ui-compat-toasts')
			?.shadowRoot?.querySelector('[data-code="SIGNATURE_STATE"]') ?? null;

	it('tells a signed deck its state and that saving removes the signatures', () => {
		act(() => root.render(<Stack signatures={deck('invalid')} />));
		expect(notice()?.getAttribute('data-severity')).toBe('warning');
		expect(notice()?.textContent).toContain('is not valid');
		expect(notice()?.textContent).toContain('Saving changes removes the signatures');
		expect(notice()?.textContent).not.toContain('This presentation is digitally signed.');
	});

	it('shows nothing for an unsigned deck', () => {
		act(() => root.render(<Stack signatures={undefined} />));
		expect(notice()).toBeNull();
	});

	it('stays dismissed until another deck loads', () => {
		act(() => root.render(<Stack signatures={deck('verified-untrusted')} />));
		const dismiss = notice()?.querySelector<HTMLButtonElement>(
			'[data-testid="pptx-compat-toast-dismiss"]',
		);
		act(() => dismiss?.click());
		expect(notice()).toBeNull();
		act(() => root.render(<Stack signatures={deck('verified-untrusted')} />));
		expect(notice()?.getAttribute('data-severity')).toBe('info');
	});

	it('is worded in the UI language', () => {
		act(() =>
			root.render(
				<PptxI18nProvider translate={() => undefined} locale='fr'>
					<Stack signatures={deck('invalid')} />
				</PptxI18nProvider>,
			),
		);
		expect(notice()?.textContent).not.toContain('Saving changes');
	});
});
