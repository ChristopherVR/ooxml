import { describe, expect, test } from 'vitest';

import type { SignatureState, SignatureSummary } from 'ooxml-core/pptx';

import { SIGNATURE_NOTICE_ID, signatureLoadNotice, signatureStatusView } from './signature-status';

const deck = (state: SignatureState): SignatureSummary => ({
	state,
	signatures: [
		{
			path: '_xmlsignatures/sig1.xml',
			state,
			signer: 'CN=Ada Lovelace, O=Analytical Engines',
			issuer: 'CN=Test CA',
			signingTime: '2026-01-02T03:04:05Z',
		},
	],
});

describe('signatureStatusView', () => {
	test('an unsigned deck shows no badge and is not signed', () => {
		const view = signatureStatusView(undefined, 'en');
		expect(view).toMatchObject({
			state: 'unsigned',
			present: false,
			signed: false,
			label: '',
			description: '',
			count: 0,
			rows: [],
		});
	});

	test.each([
		['verified', true, 'Signed', 'success', 'shield-check'],
		['verified-untrusted', true, 'Signed, signer not verified', 'caution', 'shield-alert'],
		['invalid', false, 'Invalid signature', 'danger', 'shield-x'],
		['unsupported', false, 'Signature not verified', 'neutral', 'shield-question'],
	] as const)('%s reads signed=%s as "%s"', (state, signed, label, tone, icon) => {
		const view = signatureStatusView(deck(state), 'en');
		expect(view).toMatchObject({ state, present: true, signed, label, tone, icon, count: 1 });
		expect(view.description).not.toBe('');
		expect(view.rows[0]).toMatchObject({ state, label, tone, icon, signer: 'Ada Lovelace' });
	});

	test('an invalid or unsupported signature is never labelled plain "Signed"', () => {
		for (const state of ['invalid', 'unsupported', 'verified-untrusted'] as const) {
			expect(signatureStatusView(deck(state), 'en').label).not.toBe('Signed');
		}
	});

	test('rows carry the certificate details and fall back to the part path', () => {
		const view = signatureStatusView(
			{
				state: 'invalid',
				signatures: [{ path: '_xmlsignatures/sig2.xml', state: 'invalid' }],
			},
			'en',
		);
		expect(view.rows[0]).toEqual({
			path: '_xmlsignatures/sig2.xml',
			state: 'invalid',
			label: 'Invalid signature',
			tone: 'danger',
			icon: 'shield-x',
			signer: '_xmlsignatures/sig2.xml',
		});
		const full = signatureStatusView(deck('verified-untrusted'), 'en').rows[0];
		expect(full).toMatchObject({
			subject: 'CN=Ada Lovelace, O=Analytical Engines',
			issuer: 'CN=Test CA',
		});
		expect(full?.signedAt).toBeTruthy();
	});

	test('wording follows the locale', () => {
		const view = signatureStatusView(deck('invalid'), 'fr-FR');
		expect(view.label).toBe('Signature non valide');
		expect(view.heading).not.toBe('Digital signatures');
	});
});

describe('signatureLoadNotice', () => {
	test('an unsigned deck has no notice', () => {
		expect(signatureLoadNotice(signatureStatusView(undefined, 'en'))).toBeUndefined();
	});

	test.each([
		['verified-untrusted', 'info', 'has not changed since it was signed'],
		['invalid', 'warning', 'is not valid'],
		['unsupported', 'warning', 'could not be verified here'],
	] as const)('a %s deck gets a %s notice', (state, severity, phrase) => {
		const notice = signatureLoadNotice(signatureStatusView(deck(state), 'en'));
		expect(notice).toMatchObject({ id: SIGNATURE_NOTICE_ID, severity });
		expect(notice?.text).toContain(phrase);
		expect(notice?.text).toContain('Saving changes removes the signatures');
	});

	test('the notice is worded in the locale', () => {
		const view = signatureStatusView(deck('invalid'), 'fr');
		expect(signatureLoadNotice(view)?.text).toBe(`${view.description} ${view.removedOnSave}`);
		expect(signatureLoadNotice(view)?.text).not.toContain('Saving changes');
	});
});
