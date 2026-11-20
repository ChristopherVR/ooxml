/**
 * external-content-bar.component.test.ts: the Angular adapter around the shared
 * `pptx-ui-external-content-bar` (PowerPoint's security bar for linked content).
 */
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { ExternalContentBarComponent } from './external-content-bar.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

function open() {
	TestBed.configureTestingModule({
		imports: [ExternalContentBarComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	const fixture = TestBed.createComponent(ExternalContentBarComponent);
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	fixture.detectChanges();
	const host = (fixture.nativeElement as HTMLElement).querySelector(
		'pptx-ui-external-content-bar',
	)!;
	const part = (testId: string) =>
		host.shadowRoot!.querySelector<HTMLElement>(`[data-testid="${testId}"]`)!;
	return { fixture, host, part };
}

describe('externalContentBarComponent adapter', () => {
	it("shows PowerPoint's translated security warning", () => {
		const { host, part } = open();
		expect(host.getAttribute('data-testid')).toBe('pptx-external-content-bar');
		expect(host.shadowRoot!.textContent).toContain('SECURITY WARNING');
		expect(host.shadowRoot!.textContent).toContain(
			'References to external pictures and other media have been blocked.',
		);
		expect(part('pptx-external-content-action').textContent).toBe('Enable Content');
	});

	it('forwards Enable Content and Close as separate outputs', () => {
		const { fixture, part } = open();
		const seen: string[] = [];
		fixture.componentInstance.enable.subscribe(() => seen.push('enable'));
		fixture.componentInstance.closeBar.subscribe(() => seen.push('close'));
		part('pptx-external-content-action').click();
		part('pptx-external-content-close').click();
		expect(seen).toStrictEqual(['enable', 'close']);
	});
});
