import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { signatureStrings } from 'ooxml-core/i18n';
import type { signature } from 'ooxml-core/opc';
import { registerOfficeUi } from '../index';
import type { OfficeUiSignatureStatus } from './signature-status';
import { createSignatureChrome } from './signature-chrome';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type Summary = signature.SignatureSummary;
const signer = 'CN=Ada Lovelace, O=Analytical Engines';
const summaryOf = (state: Summary['state']): Summary => ({
	state,
	signatures: [
		{ path: '_xmlsignatures/sig1.xml', state, signer, signingTime: '2026-11-20T03:04:05Z' },
	],
});

function mount(summary: Summary | undefined, variant = 'badge', locale = 'en') {
	const host = document.createElement('office-ui-signature-status') as OfficeUiSignatureStatus;
	host.variant = variant as OfficeUiSignatureStatus['variant'];
	host.locale = locale;
	host.summary = summary;
	document.body.append(host);
	return host;
}
const text = (host: HTMLElement, selector: string) =>
	host.shadowRoot!.querySelector(selector)?.textContent ?? '';

describe('office-ui-signature-status', () => {
	it('is hidden for an unsigned or absent summary', () => {
		expect(mount(undefined).hidden).toBe(true);
		const unsigned = mount({ state: 'unsigned', signatures: [] });
		expect(unsigned.hidden).toBe(true);
		expect(unsigned.dataset.state).toBe('unsigned');
	});

	it('labels only a verified signature "Signed"', () => {
		expect(text(mount(summaryOf('verified')), '.label')).toBe('Signed');
		for (const state of ['invalid', 'unsupported'] as const) {
			const host = mount(summaryOf(state));
			expect(host.hidden).toBe(false);
			expect(host.dataset.state).toBe(state);
			expect(text(host, '.label')).not.toMatch(/^Signed\b/u);
			expect(text(host, '.label')).toBe(signatureStrings.en.states[state].label);
		}
		expect(text(mount(summaryOf('verified-untrusted')), '.label')).toBe(
			'Signed, signer not verified',
		);
	});

	it('shows the state and every signer in the panel', () => {
		const host = mount(summaryOf('verified-untrusted'), 'panel');
		expect(text(host, '.heading')).toBe('Digital signatures');
		expect(text(host, '.description')).toBe(
			signatureStrings.en.states['verified-untrusted'].description,
		);
		expect(text(host, '.signer')).toBe('Ada Lovelace');
		expect(text(host, '.removed')).toBe(signatureStrings.en.removedOnSave);
	});

	it('emits view and dismiss from the load notice', () => {
		const host = mount(summaryOf('invalid'), 'notice');
		const view = vi.fn();
		const dismiss = vi.fn();
		host.addEventListener('office-signature-activate', (event) =>
			view((event as CustomEvent).detail),
		);
		host.addEventListener('office-signature-dismiss', dismiss);
		expect(text(host, '.description')).toContain(signatureStrings.en.removedOnSave);
		host.shadowRoot!.querySelector<HTMLButtonElement>('.view')!.click();
		host.shadowRoot!.querySelector<HTMLButtonElement>('.dismiss')!.click();
		expect(view).toHaveBeenCalledWith({ state: 'invalid' });
		expect(dismiss).toHaveBeenCalledTimes(1);
	});

	it.each(['fr', 'de', 'es', 'zh-CN', 'ja', 'ko'] as const)('words every state in %s', (locale) => {
		for (const state of ['verified', 'verified-untrusted', 'invalid', 'unsupported'] as const) {
			const host = mount(summaryOf(state), 'panel', locale);
			const words = signatureStrings[locale];
			expect(text(host, '.label')).toBe(words.states[state].label);
			expect(text(host, '.heading')).toBe(words.heading);
			expect(text(host, '.label')).not.toBe(signatureStrings.en.states[state].label);
		}
	});

	it('follows a locale change on the same element', () => {
		const host = mount(summaryOf('invalid'));
		host.locale = 'de-AT';
		expect(text(host, '.label')).toBe(signatureStrings.de.states.invalid.label);
	});

	it('shows the state without opening anything while the chrome is not activatable', () => {
		const open = vi.fn();
		const chrome = createSignatureChrome(document, open);
		document.body.append(chrome.badge, chrome.notice);
		chrome.update(summaryOf('invalid'), 'en');
		chrome.setActivatable(false);
		expect(chrome.badge.hasAttribute('no-activate')).toBe(true);
		expect(chrome.badge.hidden).toBe(false);
		expect(text(chrome.badge, '.label')).toBe(signatureStrings.en.states.invalid.label);
		expect(chrome.badge.shadowRoot!.querySelector('button')).toBeNull();
		expect(chrome.notice.shadowRoot!.querySelector('.view')).toBeNull();
		chrome.badge.dispatchEvent(new CustomEvent('office-signature-activate'));
		expect(open).not.toHaveBeenCalled();
		chrome.setActivatable(true);
		chrome.badge.shadowRoot!.querySelector<HTMLButtonElement>('button')!.click();
		chrome.notice.shadowRoot!.querySelector<HTMLButtonElement>('.view')!.click();
		expect(open).toHaveBeenCalledTimes(2);
	});
});
