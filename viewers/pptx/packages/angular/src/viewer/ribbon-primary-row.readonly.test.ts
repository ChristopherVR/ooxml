/**
 * The Read-only badge in the Angular quick-access row follows the shared
 * `resolveReadOnlyBadge` decision and updates when editing is allowed again.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { labelNameMismatches } from '../../../../../../src/ui/src/test-support/label-in-name';
import { translationsDe } from '../../../locales/src/de';
import { RibbonPrimaryRowComponent } from './ribbon-primary-row.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => TestBed.resetTestingModule());

function open(canEdit: boolean, lang: 'en' | 'de' = 'en') {
	TestBed.resetTestingModule();
	TestBed.configureTestingModule({
		imports: [RibbonPrimaryRowComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs: declare the one under test as a plain signal.
	TestBed.overrideComponent(RibbonPrimaryRowComponent, { add: { inputs: ['canEdit'] } });
	const fixture = TestBed.createComponent(RibbonPrimaryRowComponent);
	const editable = signal(canEdit);
	fixture.componentRef.setInput('canEdit', editable);
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.setTranslation('de', { ...translationsEn, ...translationsDe });
	translate.use(lang);
	fixture.detectChanges();
	const badge = () =>
		(fixture.nativeElement as HTMLElement).querySelector<HTMLElement>(
			'[data-pptx-read-only-badge]',
		);
	return { fixture, badge, editable };
}

describe('ribbon primary row read-only badge', () => {
	it('shows the badge for a read-only deck with a status role', () => {
		const { badge } = open(false);
		expect(badge()?.textContent?.trim()).toBe('Read-only');
		expect(badge()?.getAttribute('role')).toBe('status');
	});

	it('omits the badge when editing is allowed', () => {
		expect(open(true).badge()).toBeNull();
	});
});

describe('ribbon primary row names (WCAG 2.5.3)', () => {
	it.each([
		['en', 'Show', 'Custom shows'],
		['de', 'Zeigen', 'Benutzerdefinierte Bildschirmpräsentationen'],
	] as const)('names every labelled button with the text it shows in %s', (lang, caption, tip) => {
		const { fixture } = open(true, lang);
		const root = fixture.nativeElement as HTMLElement;
		const pill = [...root.querySelectorAll<HTMLButtonElement>('.pptx-rb-pill')].find(
			(button) => button.title === tip,
		);
		expect(pill?.textContent?.trim()).toBe(caption);
		expect(labelNameMismatches(root)).toEqual([]);
	});
});
