/**
 * A deck reads as signed only when its signature verifies: the status bar badge
 * and the signatures panel show the verified state from the shared view, never
 * "Signed" for a signature part that failed or could not be checked.
 */
import { signal } from '@angular/core';
import type { Type } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import type { ComponentFixture } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureState, SignatureSummary } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { translationsFr } from '../../../locales/src';
import { SignatureStatusBadgeComponent } from './signature-status-badge.component';
import { SignaturesPanelComponent } from './signatures-panel.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => TestBed.resetTestingModule());

function summary(state: SignatureState): SignatureSummary {
	return {
		state,
		signatures: [
			{
				path: '/_xmlsignatures/sig1.xml',
				state,
				signer: 'CN=Ada Lovelace, O=Analytical',
				issuer: 'CN=Example CA',
				signingTime: '2026-03-04T05:06:07Z',
			},
		],
	};
}

function mount<T>(
	component: Type<T>,
	value: SignatureSummary | undefined,
	lang = 'en',
): ComponentFixture<T> {
	TestBed.configureTestingModule({
		imports: [component],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// JIT runs without the signal-input transform: hand the input over as a signal.
	TestBed.overrideComponent(component, { add: { inputs: ['summary'] } });
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.setTranslation('fr', translationsFr);
	translate.use(lang);
	const fixture = TestBed.createComponent(component);
	fixture.componentRef.setInput('summary', signal(value));
	fixture.detectChanges();
	return fixture;
}

const render = (value: SignatureSummary | undefined, lang = 'en'): HTMLElement =>
	mount(SignaturesPanelComponent, value, lang).nativeElement as HTMLElement;

const text = (root: HTMLElement, selector: string) =>
	root.querySelector(selector)?.textContent?.trim();

describe('signatures panel state', () => {
	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
		['verified-untrusted', 'Signed, signer not verified'],
		['verified', 'Signed'],
	] as const)('labels a %s deck "%s"', (state, label) => {
		const root = render(summary(state));
		expect(root.querySelector('section')?.getAttribute('data-pptx-signature-state')).toBe(state);
		expect(text(root, '.pptx-ng-signatures__title')).toBe(label);
		const row = root.querySelector('li');
		expect(row?.getAttribute('data-pptx-signature-row-state')).toBe(state);
		expect(text(row as HTMLElement, '.pptx-ng-signatures__badge')).toBe(label);
		expect(text(row as HTMLElement, '.pptx-ng-signatures__signer')).toBe('Ada Lovelace');
	});

	it('never reads "Signed" for a signature that did not verify', () => {
		for (const state of ['invalid', 'unsupported'] as const) {
			const root = render(summary(state));
			expect(root.textContent).not.toMatch(/\bSigned\b/);
			TestBed.resetTestingModule();
		}
	});

	it('shows the issuer, signing time caption and the removed-on-save note', () => {
		const root = render(summary('verified-untrusted'));
		const meta = root.querySelector('dl')?.textContent ?? '';
		expect(meta).toContain('Issuer');
		expect(meta).toContain('CN=Example CA');
		expect(meta).toContain('Signed on');
		expect(text(root, '.pptx-ng-signatures__note')).toBe(
			'Saving changes removes the signatures, because any change invalidates them.',
		);
	});

	it('shows the no-signatures message for an unsigned deck', () => {
		const root = render(undefined);
		expect(root.querySelector('section')?.getAttribute('data-pptx-signature-state')).toBe(
			'unsigned',
		);
		expect(text(root, '.pptx-ng-signatures__empty')).toBe(
			translationsEn['pptx.digitalSignatures.noSignatures'],
		);
		expect(root.querySelector('li')).toBeNull();
	});

	it('follows the viewer language', () => {
		const root = render(summary('invalid'), 'fr');
		expect(text(root, '.pptx-ng-signatures__title')).toBe('Signature non valide');
	});
});

describe('signature status badge', () => {
	it.each([
		['invalid', 'Invalid signature'],
		['unsupported', 'Signature not verified'],
		['verified-untrusted', 'Signed, signer not verified'],
		['verified', 'Signed'],
	] as const)('reads "%2$s" for a %1$s deck', (state, label) => {
		const root = mount(SignatureStatusBadgeComponent, summary(state)).nativeElement as HTMLElement;
		const badge = root.querySelector('[data-pptx-signature-state]');
		expect(badge?.getAttribute('data-pptx-signature-state')).toBe(state);
		expect(badge?.textContent?.trim()).toBe(label);
		expect(badge?.getAttribute('title')).toBe(
			signatureStatusView(summary(state), 'en').description,
		);
	});

	it('is absent for an unsigned deck', () => {
		const root = mount(SignatureStatusBadgeComponent, undefined).nativeElement as HTMLElement;
		expect(root.querySelector('[data-pptx-signature-state]')).toBeNull();
	});

	it('opens the signatures panel when clicked', () => {
		const fixture = mount(SignatureStatusBadgeComponent, summary('invalid'));
		let opened = 0;
		fixture.componentInstance.open.subscribe(() => opened++);
		(fixture.nativeElement as HTMLElement).querySelector('button')?.click();
		expect(opened).toBe(1);
	});
});
