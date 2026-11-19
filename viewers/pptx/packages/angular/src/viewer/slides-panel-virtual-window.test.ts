import { readFileSync } from 'node:fs';

import { NgStyle } from '@angular/common';
import { NO_ERRORS_SCHEMA, signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslatePipe, TranslateService } from '@ngx-translate/core';
import { editorThumbnailStep } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, expect, test } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { EditorStateService } from './editor-state.service';
import { SlidesPanelComponent } from './slides-panel.component';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

test('follows scrolling in the next frame, building the rows that entered', async () => {
	TestBed.configureTestingModule({
		imports: [SlidesPanelComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' }), EditorStateService],
	});
	TestBed.overrideComponent(SlidesPanelComponent, {
		set: {
			templateUrl: '',
			template: readFileSync(`${import.meta.dirname}/slides-panel.component.html`, 'utf8'),
			styleUrl: '',
			styles: [readFileSync(`${import.meta.dirname}/slides-panel.component.css`, 'utf8')],
			imports: [NgStyle, TranslatePipe],
			schemas: [NO_ERRORS_SCHEMA],
		},
	});
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
	TestBed.inject(EditorStateService).setSlides(
		Array.from(
			{ length: 100 },
			(_, i) => ({ id: `s${i}`, slideNumber: i + 1, elements: [] }) as PptxSlide,
		),
	);
	const fixture = TestBed.createComponent(SlidesPanelComponent);
	// Plain JIT does not discover signal-input metadata; supply the inputs directly.
	Object.assign(fixture.componentInstance, {
		canvasSize: signal({ width: 960, height: 540 }),
		activeIndex: signal(0),
	});
	fixture.detectChanges();
	const root = fixture.nativeElement as HTMLElement;
	const row = (n: number) => root.querySelector(`[aria-label="Go to slide ${n}"]`);
	expect(row(61)).toBeNull();
	const list = root.querySelector<HTMLElement>('[data-pptx-chrome="slide-list"]')!;
	list.scrollTop = 59 * editorThumbnailStep(960, 540);
	list.dispatchEvent(new Event('scroll'));
	await new Promise<void>((resolve) => {
		requestAnimationFrame(() => resolve());
	});
	fixture.detectChanges();
	expect(row(61)).not.toBeNull();
	expect(row(1)).toBeNull();
	fixture.destroy();
});
