/**
 * export-progress-modal.component.test.ts: the Angular adapter around the shared
 * `pptx-ui-export-progress`.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { ExportProgressModalComponent } from './export-progress-modal.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

function open(inputs: Record<string, unknown>) {
	TestBed.resetTestingModule();
	TestBed.configureTestingModule({
		imports: [ExportProgressModalComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(ExportProgressModalComponent, {
		add: { inputs: Object.keys(inputs) },
	});
	const fixture = TestBed.createComponent(ExportProgressModalComponent);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	fixture.detectChanges();
	const host = (fixture.nativeElement as HTMLElement).querySelector('pptx-ui-export-progress')!;
	const inner = (selector: string) => host.shadowRoot!.querySelector<HTMLElement>(selector);
	return { fixture, inner };
}

describe('exportProgressModalComponent adapter', () => {
	it('maps the export onto the shared progress dialog', () => {
		const { inner } = open({
			open: true,
			title: 'Export as PDF',
			progress: 40,
			statusMessage: 'Rendering slide 2 of 5...',
		});
		expect(inner('[role="dialog"]')?.getAttribute('aria-modal')).toBe('true');
		expect(inner('#title')?.textContent).toBe('Export as PDF');
		expect(inner('[role="progressbar"]')?.getAttribute('aria-valuenow')).toBe('40');
		expect(inner('#status')?.textContent).toBe('Rendering slide 2 of 5...');
		expect(inner('[data-testid="pptx-export-progress-cancel"]')?.textContent).toBe('Cancel');
	});

	it('emits cancel from Cancel, says Processing without a status, and is empty while closed', () => {
		const shown = open({ open: true, title: 'Export', progress: 0 });
		const seen: string[] = [];
		shown.fixture.componentInstance.cancel.subscribe(() => seen.push('cancel'));
		expect(shown.inner('#status')?.textContent).toBe('Processing…');
		shown.inner('[data-testid="pptx-export-progress-cancel"]')!.click();
		expect(seen).toStrictEqual(['cancel']);
		const hidden = open({ open: false, title: 'Export', progress: 0 });
		expect(hidden.inner('[role="dialog"]')).toBeNull();
	});
});
