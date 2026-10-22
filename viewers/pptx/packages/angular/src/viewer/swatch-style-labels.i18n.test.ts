/**
 * The theme colour grid and the shape panel's quick styles name their swatches through
 * ngx-translate (tooltip and accessible name), not in hard-coded English, and the quick styles
 * follow a language change. The grid is built in an injection context with stubbed services; the
 * shape panel is mounted with TestBed, since the shared gallery draws its tiles once rendered.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { ThemeColorSwatchColumn } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { translationsDe } from 'pptx-viewer-locales';
import { Subject } from 'rxjs';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { LoadContentService } from './load-content.service';
import { ShapeAuthoringPanelComponent } from './shape-authoring-panel.component';
import { ThemeColorSwatchGridComponent } from './theme-color-swatch-grid.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

const OFFICE_THEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	bg1: '#FFFFFF',
	tx1: '#000000',
	bg2: '#E7E6E6',
	tx2: '#44546A',
};

/** `TranslateService.instant` over the German dictionary, with ngx-translate's `{{name}}` filling. */
const translation = {
	onLangChange: new Subject<unknown>(),
	onTranslationChange: new Subject<unknown>(),
	instant: (key: string, params?: Record<string, string | number>) =>
		((translationsDe as Record<string, string>)[key] ?? key).replace(
			/\{\{\s*(\w+)\s*\}\}/gu,
			(match, name: string) => (params?.[name] === undefined ? match : String(params[name])),
		),
};

function inject<T>(create: () => T): T {
	const injector = Injector.create({
		providers: [
			{ provide: TranslateService, useValue: translation },
			{ provide: LoadContentService, useValue: { themeColorMap: signal(OFFICE_THEME) } },
		],
	});
	return runInInjectionContext(injector, create);
}

describe('swatch and style names follow the locale', () => {
	it('names theme colour swatches with the host dictionary', () => {
		const grid = inject(() => new ThemeColorSwatchGridComponent()) as unknown as {
			columns(): readonly ThemeColorSwatchColumn[];
			heading(): string;
		};
		const labels = grid
			.columns()
			.flatMap((column) => [column.base, ...column.variants])
			.map((swatch) => swatch.label);
		expect(labels).toContain('Akzent 1, heller 80 %');
		expect(labels.some((label) => label.includes('Accent'))).toBeFalsy();
		expect(grid.heading()).toBe('Designfarben');
	});

	it('names quick style presets with the host dictionary, and relabels them on a switch', () => {
		TestBed.configureTestingModule({
			imports: [ShapeAuthoringPanelComponent],
			providers: [provideTranslateService({ fallbackLang: 'en' })],
		});
		// The Vitest JIT build does not wire signal inputs: declare the input, hand it a signal.
		TestBed.overrideComponent(ShapeAuthoringPanelComponent, { add: { inputs: ['element'] } });
		const translate = TestBed.inject(TranslateService);
		translate.setTranslation('en', translationsEn);
		translate.setTranslation('de', translationsDe);
		translate.use('de');
		const fixture = TestBed.createComponent(ShapeAuthoringPanelComponent);
		fixture.componentRef.setInput(
			'element',
			signal({ id: 'box', type: 'shape', x: 0, y: 0, width: 80, height: 40 } as PptxElement),
		);
		fixture.detectChanges();
		const tile = () =>
			(fixture.nativeElement as HTMLElement).querySelector<HTMLButtonElement>(
				'[data-gallery-item]',
			)!;
		expect(tile().getAttribute('aria-label')).toBe('Blaue Füllung');
		expect(tile().title).toBe('Blaue Füllung');

		translate.use('en');
		fixture.detectChanges();
		expect(tile().getAttribute('aria-label')).toBe('Blue Fill');
		expect(tile().title).toBe('Blue Fill');
	});
});
