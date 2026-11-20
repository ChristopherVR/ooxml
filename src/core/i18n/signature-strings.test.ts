import { describe, expect, it } from 'vitest';

import { SIGNATURE_STATES } from '../opc/signature/state';
import { signatureStateText, signatureStrings, signatureStringsFor } from './signature-strings';

describe('signatureStrings', () => {
	it('words every signed state in every locale, never as English outside English', () => {
		const english = signatureStrings.en;
		for (const [locale, table] of Object.entries(signatureStrings)) {
			for (const state of SIGNATURE_STATES.filter((value) => value !== 'unsigned')) {
				const text = table.states[state as keyof typeof table.states];
				expect(text.label.length).toBeGreaterThan(0);
				expect(text.description.length).toBeGreaterThan(text.label.length);
				if (locale !== 'en') {
					expect(text.label).not.toBe(english.states[state as keyof typeof english.states].label);
				}
			}
			if (locale === 'en') continue;
			for (const key of ['removedOnSave', 'heading', 'signer', 'signedAt', 'view'] as const)
				expect(table[key], `${locale} ${key}`).not.toBe(english[key]);
			expect(table.dismiss.length).toBeGreaterThan(0);
		}
	});

	it('calls only a verified signature "Signed" with no qualifier', () => {
		const { states } = signatureStrings.en;
		expect(states.verified.label).toBe('Signed');
		for (const state of ['invalid', 'unsupported'] as const)
			expect(states[state].label).not.toMatch(/^Signed\b/u);
		expect(states['verified-untrusted'].label).toMatch(/not verified/u);
	});

	it('shows nothing for an unsigned document and resolves locale tags like the products', () => {
		expect(signatureStateText('unsigned', 'de')).toBeUndefined();
		expect(signatureStateText('invalid', 'de-AT')).toBe(signatureStrings.de.states.invalid);
		expect(signatureStringsFor('zh')).toBe(signatureStrings['zh-CN']);
		expect(signatureStringsFor('pt-BR')).toBe(signatureStrings.en);
	});
});
