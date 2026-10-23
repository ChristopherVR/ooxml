import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { signal, ɵresolveComponentResources as resolveComponentResources } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import type { ChartPptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { AdvancedChartEditorComponent } from './advanced-chart-editor.component';
import { ChartDataEditorComponent } from './chart-data-editor.component';
import { ChartFilteredSeriesOptionsComponent } from './chart-filtered-series-options.component';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
	// The JIT build leaves `templateUrl`/`styleUrl` unresolved; read them from disk.
	await resolveComponentResources(async (url) =>
		readFileSync(resolve(import.meta.dirname, url), 'utf8'),
	);
});
afterEach(() => TestBed.resetTestingModule());

const chart = {
	id: 'c1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 100,
	height: 100,
	chartData: {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [{ name: 'North', values: [1, 2], color: '#112233' }],
	},
} as ChartPptxElement;

function mount(inputs: Record<string, unknown>) {
	TestBed.configureTestingModule({
		imports: [ChartDataEditorComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	// The nested advanced editors are covered by their own tests; leave them as inert tags here.
	TestBed.overrideComponent(ChartDataEditorComponent, {
		remove: { imports: [AdvancedChartEditorComponent, ChartFilteredSeriesOptionsComponent] },
		add: { inputs: Object.keys(inputs) },
	});
	const fixture = TestBed.createComponent(ChartDataEditorComponent);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	const emitted: ChartPptxElement[] = [];
	fixture.componentInstance.elementChange.subscribe((next) => emitted.push(next));
	fixture.detectChanges();
	const sheet = (fixture.nativeElement as HTMLElement).querySelector('pptx-ui-chart-data-sheet')!;
	const cell = (row: number, col: number) =>
		sheet.shadowRoot!.querySelector<HTMLInputElement>(
			`input[data-row="${row}"][data-col="${col}"]`,
		)!;
	return { fixture, sheet, cell, emitted };
}

describe('chartDataEditorComponent (angular)', () => {
	it('feeds the shared sheet and emits the edited chart as a whole element', () => {
		const { cell, emitted } = mount({ element: chart, canEdit: true });
		expect(cell(2, 1).value).toBe('2');
		cell(2, 1).value = '7';
		cell(2, 1).dispatchEvent(new Event('change', { bubbles: true }));
		expect(emitted).toHaveLength(1);
		expect(emitted[0]!.id).toBe('c1');
		expect(emitted[0]!.chartData!.series[0]).toMatchObject({ values: [1, 7], color: '#112233' });
	});

	it('is read-only without edit rights', () => {
		const { sheet, cell } = mount({ element: chart, canEdit: false });
		expect(cell(1, 1).disabled).toBeTruthy();
		expect(sheet.shadowRoot!.querySelector('[data-action="add-series"]')).toBeNull();
	});
});
