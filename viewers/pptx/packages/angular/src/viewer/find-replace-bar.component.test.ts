/**
 * The Find and Replace bar is presentational: its query, replacement and scopes come from the
 * viewer's session (so a reopened bar shows what the session kept), and every edit goes back out.
 *
 * Reference binding: packages/react/src/viewer/components/FindReplacePanel.tsx
 */
import { signal } from '@angular/core';
import type { WritableSignal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n';
import { FindReplaceBarComponent } from './find-replace-bar.component';
import type { FindEvent, ReplaceEvent } from './find-replace-bar.component';

beforeAll(() => {
	try {
		TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	} catch {
		// Already initialised by another spec in this worker.
	}
});
afterEach(() => {
	TestBed.resetTestingModule();
	document.body.replaceChildren();
});

function open(values: Record<string, unknown>) {
	TestBed.resetTestingModule();
	TestBed.configureTestingModule({
		imports: [FindReplaceBarComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(FindReplaceBarComponent, { add: { inputs: Object.keys(values) } });
	const fixture = TestBed.createComponent(FindReplaceBarComponent);
	for (const [name, value] of Object.entries(values)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.use('en');
	fixture.detectChanges();
	return fixture;
}

const KEPT = {
	query: 'quarter',
	replacement: 'period',
	matchCase: true,
	includeNotes: false,
	includeChartTitles: true,
	includeChartData: false,
	matchCount: 3,
	matchIndex: 1,
};

describe('findReplaceBarComponent', () => {
	it('shows the query, replacement and status the session kept', () => {
		const root = open(KEPT).nativeElement as HTMLElement;
		const [query, replacement] = Array.from(root.querySelectorAll('input'));
		expect(query?.value).toBe('quarter');
		expect(replacement?.value).toBe('period');
		expect(root.querySelector('.pptx-frb__count')?.textContent?.trim()).toBe('2 of 3');
	});

	it('sends a typed replacement back out', () => {
		const fixture = open(KEPT);
		const seen: string[] = [];
		fixture.componentInstance.replacementChange.subscribe((value) => seen.push(value));
		const replacement = (fixture.nativeElement as HTMLElement).querySelectorAll('input')[1]!;
		replacement.value = 'season';
		replacement.dispatchEvent(new Event('input'));
		expect(seen).toStrictEqual(['season']);
	});

	it('a scope toggle searches with the kept query and the other scopes as shown', () => {
		const fixture = open(KEPT);
		const seen: FindEvent[] = [];
		fixture.componentInstance.find.subscribe((event) => seen.push(event));
		fixture.componentInstance.toggleCase();
		fixture.componentInstance.toggleIncludeNotes();
		expect(seen).toStrictEqual([
			{
				query: 'quarter',
				matchCase: false,
				includeNotes: false,
				includeChartTitles: true,
				includeChartData: false,
			},
			{
				query: 'quarter',
				matchCase: false,
				includeNotes: true,
				includeChartTitles: true,
				includeChartData: false,
			},
		]);
	});

	// Zoneless change detection hands the bar the session's values on its next pass, so a
	// second edit can arrive while the inputs still show the session from before the first.
	it('an edit made before the session feeds back keeps the edit before it', () => {
		const fixture = open(KEPT);
		const root = fixture.nativeElement as HTMLElement;
		const seen: FindEvent[] = [];
		fixture.componentInstance.find.subscribe((event) => seen.push(event));
		root
			.querySelector('pptx-ui-checkbox[aria-label="Include chart series and category labels"]')!
			.dispatchEvent(new Event('change'));
		const query = root.querySelectorAll('input')[0]!;
		query.value = 'Dataset';
		query.dispatchEvent(new Event('input'));
		expect(seen.map((event) => [event.query, event.includeChartData])).toStrictEqual([
			['quarter', true],
			['Dataset', true],
		]);
	});

	it('what the session hands back replaces what the bar shows', () => {
		const fixture = open(KEPT);
		const query = (fixture.nativeElement as HTMLElement).querySelectorAll('input')[0]!;
		query.value = 'quart';
		query.dispatchEvent(new Event('input'));
		(fixture.componentInstance.query as unknown as WritableSignal<string>).set('quarterly');
		fixture.detectChanges();
		expect(query.value).toBe('quarterly');
	});

	it('replace sends the replacement just typed', () => {
		const fixture = open(KEPT);
		const seen: ReplaceEvent[] = [];
		fixture.componentInstance.replaceOne.subscribe((event) => seen.push(event));
		const replacement = (fixture.nativeElement as HTMLElement).querySelectorAll('input')[1]!;
		replacement.value = 'season';
		replacement.dispatchEvent(new Event('input'));
		replacement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(seen.map((event) => event.replacement)).toStrictEqual(['season']);
	});
});
