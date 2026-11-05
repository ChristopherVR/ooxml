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
import { RibbonPrimaryRowComponent } from './ribbon-primary-row.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
});
afterEach(() => TestBed.resetTestingModule());

function open(canEdit: boolean) {
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
	translate.use('en');
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
