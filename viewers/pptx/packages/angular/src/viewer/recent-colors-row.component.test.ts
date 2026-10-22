/**
 * recent-colors-row.component.test.ts: the reusable "Recent colours" row
 * (wave-4 B6) under the inspector's fill, stroke and text colour pickers,
 * drawn by the shared `pptx-ui-color-grid` with only Recent Colors.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { RecentColorsRowComponent } from './recent-colors-row.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

function mount(inputs: Record<string, unknown>) {
	TestBed.configureTestingModule({
		imports: [RecentColorsRowComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(RecentColorsRowComponent, {
		add: { inputs: Object.keys(inputs) },
	});
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', {
		'pptx.colorPicker.recentColors': 'Recent Colors',
		'pptx.colorPicker.recentColor': 'Recent {{hex}}',
	});
	translate.setTranslation('de', {
		'pptx.colorPicker.recentColors': 'Zuletzt verwendete Farben',
		'pptx.colorPicker.recentColor': 'Zuletzt verwendete Farbe {{hex}}',
	});
	translate.use('en');
	const fixture = TestBed.createComponent(RecentColorsRowComponent);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	fixture.detectChanges();
	const row = () =>
		(fixture.nativeElement as HTMLElement).querySelector<HTMLElement>(
			'[data-testid="pptx-color-recent"]',
		);
	const swatches = () => [
		...(row()?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? []),
	];
	const picked: string[] = [];
	fixture.componentInstance.pick.subscribe((hex) => picked.push(hex));
	return { fixture, translate, row, swatches, picked };
}

describe('recentColorsRowComponent', () => {
	it('renders nothing while the list is empty', () => {
		expect(mount({ colors: [] }).row()).toBeNull();
	});

	it('is the shared grid with only Recent Colors, one named swatch per colour', () => {
		const { row, swatches } = mount({ colors: ['#112233', '#445566'] });
		expect(row()?.localName).toBe('pptx-ui-color-grid');
		expect(row()?.getAttribute('aria-label')).toBe('Recent Colors');
		expect(
			[...row()!.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toStrictEqual(['Recent Colors']);
		expect(swatches().map((swatch) => swatch.getAttribute('aria-label'))).toStrictEqual([
			'Recent #112233',
			'Recent #445566',
		]);
	});

	it('emits the picked colour', () => {
		const { swatches, picked } = mount({ colors: ['#112233', '#445566'] });
		swatches()[1]!.click();
		expect(picked).toStrictEqual(['#445566']);
	});

	it('disables every swatch while the picker is disabled', () => {
		const { swatches } = mount({ colors: ['#112233'], disabled: true });
		expect(swatches().every((swatch) => swatch.disabled)).toBeTruthy();
	});

	it('follows a language switch', () => {
		const { fixture, translate, row, swatches } = mount({ colors: ['#112233'] });
		translate.use('de');
		fixture.detectChanges();
		expect(row()?.getAttribute('aria-label')).toBe('Zuletzt verwendete Farben');
		expect(swatches()[0]!.getAttribute('aria-label')).toBe('Zuletzt verwendete Farbe #112233');
	});
});
