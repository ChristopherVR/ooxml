/**
 * The theme colour grid and the shape panel's quick styles name their swatches through
 * ngx-translate (tooltip and accessible name), not in hard-coded English. No TestBed (see
 * `vitest.config.ts`): the components are built in an injection context with stubbed services.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import type { ShapeQuickStyle, ThemeColorSwatchColumn } from 'ooxml-ui/pptx';
import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import { translationsDe } from 'pptx-viewer-locales';
import { Subject } from 'rxjs';
import { describe, expect, it } from 'vitest';

import { LoadContentService } from './load-content.service';
import { ShapeAuthoringPanelComponent } from './shape-authoring-panel.component';
import { ThemeColorSwatchGridComponent } from './theme-color-swatch-grid.component';

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
		};
		const labels = grid
			.columns()
			.flatMap((column) => [column.base, ...column.variants])
			.map((swatch) => swatch.label);
		expect(labels).toContain('Akzent 1, heller 80 %');
		expect(labels.some((label) => label.includes('Accent'))).toBeFalsy();
	});

	it('names quick style presets with the host dictionary', () => {
		const panel = inject(() => new ShapeAuthoringPanelComponent()) as unknown as {
			quickStyleLabel(style: ShapeQuickStyle): string;
		};
		expect(panel.quickStyleLabel(SHAPE_QUICK_STYLES[0]!)).toBe('Blaue Füllung');
	});
});
