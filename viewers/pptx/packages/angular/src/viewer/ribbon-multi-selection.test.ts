import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService } from '@ngx-translate/core';
import { ribbonTab } from 'ooxml-ui/pptx';
import { afterEach, beforeAll, expect, test, vi } from 'vitest';

import { resolveViewerComponentResources } from './component-resources.test-support';
import { PowerPointViewerComponent } from './power-point-viewer.component';
import { declareSignalFields } from './signal-inputs.test-support';

beforeAll(async () => {
	declareSignalFields(PowerPointViewerComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

test('keeps Shape Format and Merge Shapes available when selecting a second shape', async () => {
	TestBed.configureTestingModule({ providers: [provideTranslateService()] });
	const fixture = TestBed.createComponent(PowerPointViewerComponent);
	const viewer = fixture.componentInstance;
	const root: HTMLElement = fixture.nativeElement;
	const deck = resolve(import.meta.dirname, '../../../../../../e2e/pptx/fixtures/merge-crop.pptx');
	fixture.componentRef.setInput('content', new Uint8Array(readFileSync(deck)));
	fixture.componentRef.setInput('canEditInput', true);
	fixture.componentRef.setInput('googleFonts', false);
	await vi.waitFor(
		() => {
			fixture.detectChanges();
			expect(viewer.getElements(0)).toHaveLength(3);
		},
		{ timeout: 5000 },
	);
	const shapes = viewer
		.getElements(0)
		.filter(({ type }) => type === 'shape')
		.map(({ id }) => id);
	expect(shapes).toHaveLength(2);
	const tab = () => ribbonTab(root, '[data-ribbon-contextual-tab="shapeFormat"]');
	const merge = () =>
		root.querySelector(
			'[data-ribbon-group="shapeFormat.arrange"] [data-pptx-ribbon-control="merge-shapes"]',
		);
	viewer.selectElements(shapes.slice(0, 1));
	fixture.detectChanges();
	tab()!.click();
	fixture.detectChanges();
	expect(merge()).toHaveProperty('disabled', true);
	viewer.selectElements(shapes);
	fixture.detectChanges();
	expect(tab()?.getAttribute('aria-selected')).toBe('true');
	expect(merge()).toHaveProperty('disabled', false);
});
