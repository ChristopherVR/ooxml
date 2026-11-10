/**
 * A thumbnail's chart subscribes to the document-wide picture colour sampler;
 * destroying the thumbnail (a host strip dropping or replacing it) must retire
 * that subscription, not leave it for the next sample to find.
 */
import { Component, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { SlideThumbnailComponent } from '../public-api';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { declareSignalFields } from './signal-inputs.test-support';

const sampler = vi.hoisted(() => ({ active: 0 }));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		subscribeBarFacePicturePixelSamples: (listener: () => void) => {
			sampler.active += 1;
			const off = actual.subscribeBarFacePicturePixelSamples(listener);
			let open = true;
			return () => {
				if (open) {
					open = false;
					sampler.active -= 1;
				}
				off();
			};
		},
	};
});

const chart = {
	type: 'chart',
	id: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: { chartType: 'bar', categories: ['A'], series: [{ name: 'One', values: [1] }] },
} as unknown as PptxElement;

@Component({
	selector: 'test-chart-thumbnail-host',
	standalone: true,
	imports: [SlideThumbnailComponent],
	template: `<pptx-slide-thumbnail [slide]="slide()" [canvasSize]="{ width: 960, height: 540 }" />`,
})
class ChartThumbnailHostComponent {
	readonly slide = signal<PptxSlide>({
		id: 'ppt/slides/slide1.xml',
		rId: 'rId2',
		slideNumber: 1,
		elements: [chart],
	} as unknown as PptxSlide);
}

beforeAll(async () => {
	declareSignalFields(ChartThumbnailHostComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

describe('slideThumbnailComponent chart subscription', () => {
	it('retires the chart sampler subscription when the thumbnail is destroyed', () => {
		TestBed.configureTestingModule({
			providers: [provideTranslateService({ fallbackLang: 'en' })],
		});
		TestBed.inject(TranslateService).setTranslation('en', translationsEn);
		TestBed.inject(TranslateService).use('en');
		const fixture = TestBed.createComponent(ChartThumbnailHostComponent);
		fixture.detectChanges();
		expect((fixture.nativeElement as HTMLElement).querySelector('svg')).not.toBeNull();
		expect(sampler.active).toBe(1);
		fixture.destroy();
		expect(sampler.active).toBe(0);
	});
});
