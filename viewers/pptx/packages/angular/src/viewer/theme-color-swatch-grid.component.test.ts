import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { LoadContentService } from './load-content.service';
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

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

function mount(themeColorMap: Record<string, string> | undefined, inputs: Record<string, unknown>) {
	TestBed.configureTestingModule({
		imports: [ThemeColorSwatchGridComponent],
		providers: [
			provideTranslateService({ fallbackLang: 'en' }),
			{ provide: LoadContentService, useValue: { themeColorMap: signal(themeColorMap) } },
		],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(ThemeColorSwatchGridComponent, {
		add: { inputs: Object.keys(inputs) },
	});
	const fixture = TestBed.createComponent(ThemeColorSwatchGridComponent);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	fixture.detectChanges();
	const grid = (fixture.nativeElement as HTMLElement).querySelector('pptx-ui-color-grid');
	const swatch = (title: string) =>
		grid?.shadowRoot?.querySelector<HTMLButtonElement>(`button[title="${title}"]`) ?? null;
	return { fixture, grid, swatch };
}

describe('themeColorSwatchGridComponent', () => {
	it('renders nothing when no deck theme is loaded', () => {
		expect(mount(undefined, {}).grid).toBeNull();
	});

	it('draws the deck theme on the shared grid without a standard row', () => {
		const { grid } = mount(OFFICE_THEME, {});
		const shadow = grid!.shadowRoot!;
		// 10 columns x (1 base + 5 variants) = 60 swatch buttons.
		expect(shadow.querySelectorAll('button.swatch[data-source="theme"]')).toHaveLength(60);
		expect(shadow.querySelectorAll('button.swatch')).toHaveLength(60);
		expect(shadow.querySelectorAll('.heading')).toHaveLength(1);
	});

	it('marks the swatch the stored theme reference selects', () => {
		const { grid } = mount(OFFICE_THEME, {
			selectedRef: { scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 },
			selectedHex: '#000000',
		});
		const checked = grid!.shadowRoot!.querySelectorAll('button[aria-checked="true"]');
		expect([...checked].map((el) => el.getAttribute('title'))).toStrictEqual([
			'Accent 1, Lighter 80%',
		]);
	});

	it('emits both the hex and the ref for a picked swatch', () => {
		const { fixture, swatch } = mount(OFFICE_THEME, {});
		const picks: ThemeColorPickerCommit[] = [];
		fixture.componentInstance.pick.subscribe((commit) => picks.push(commit));
		swatch('Accent 1')!.click();
		expect(picks).toStrictEqual([{ hex: '#4472c4', ref: { scheme: 'accent1' } }]);
	});

	it('disables every swatch when the field is disabled', () => {
		const { grid } = mount(OFFICE_THEME, { disabled: true });
		const swatches = [...grid!.shadowRoot!.querySelectorAll<HTMLButtonElement>('button.swatch')];
		expect(swatches.every((button) => button.disabled)).toBeTruthy();
	});
});
