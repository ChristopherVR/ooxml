/**
 * `SlideThumbnailComponent`: one slide drawn the way the rail draws it, scaled
 * to the width a host picks, template elements beneath the slide's own, and
 * redrawn when the slide model changes.
 */
import { Component, signal } from '@angular/core';
import type { ComponentFixture } from '@angular/core/testing';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
// Import through the package entry so the test also covers the public export.
import { SlideThumbnailComponent } from '../public-api';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { declareSignalFields } from './signal-inputs.test-support';

function shape(id: string, x = 0): PptxElement {
	return {
		id,
		type: 'shape',
		shapeType: 'rect',
		x,
		y: 0,
		width: 100,
		height: 50,
		shapeStyle: { fillColor: '#ff0000' },
	} as unknown as PptxElement;
}

function slide(elements: PptxElement[]): PptxSlide {
	return {
		id: 'ppt/slides/slide1.xml',
		rId: 'rId2',
		slideNumber: 1,
		elements,
		showMasterShapes: true,
	} as unknown as PptxSlide;
}

@Component({
	selector: 'test-thumbnail-host',
	standalone: true,
	imports: [SlideThumbnailComponent],
	template: `
		<pptx-slide-thumbnail
			[slide]="slide()"
			[templateElements]="templates()"
			[canvasSize]="{ width: 960, height: 540 }"
			[previewWidth]="120"
		/>
	`,
})
class ThumbnailHostComponent {
	readonly slide = signal<PptxSlide>(slide([]));
	readonly templates = signal<readonly PptxElement[]>([]);
}

beforeAll(async () => {
	declareSignalFields(ThumbnailHostComponent, import.meta.dirname);
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => TestBed.resetTestingModule());

function mount(
	elements: PptxElement[],
	templates: PptxElement[] = [],
): ComponentFixture<ThumbnailHostComponent> {
	TestBed.configureTestingModule({ providers: [provideTranslateService({ fallbackLang: 'en' })] });
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
	const fixture = TestBed.createComponent(ThumbnailHostComponent);
	fixture.componentInstance.slide.set(slide(elements));
	fixture.componentInstance.templates.set(templates);
	fixture.detectChanges();
	return fixture;
}

const root = (fixture: ComponentFixture<unknown>): HTMLElement =>
	fixture.nativeElement as HTMLElement;
const box = (fixture: ComponentFixture<unknown>) =>
	root(fixture).querySelector<HTMLElement>('.pptx-ng-slide-thumbnail');
const stage = (fixture: ComponentFixture<unknown>) =>
	root(fixture).querySelector<HTMLElement>('.pptx-ng-canvas-stage');
const drawn = (fixture: ComponentFixture<unknown>) => [
	...(stage(fixture)?.querySelectorAll<HTMLElement>('.pptx-ng-element') ?? []),
];

describe('slideThumbnailComponent', () => {
	it('draws the slide stage scaled to the width the host picks, at the deck aspect ratio', () => {
		const fixture = mount([shape('a')]);
		expect(box(fixture)?.style.height).toBe('67.5px');
		expect(stage(fixture)?.style.width).toBe('960px');
		expect(stage(fixture)?.style.transform).toBe('scale(0.125)');
	});

	it('draws inherited template elements beneath the slide elements', () => {
		const fixture = mount([shape('own', 200)], [shape('layout', 10)]);
		expect(drawn(fixture).map((node) => node.style.left)).toStrictEqual(['10px', '200px']);
	});

	it('stays static: no element ids, editing contract or slide region', () => {
		const fixture = mount([shape('a')], [shape('layout', 10)]);
		const host = root(fixture);
		expect(host.querySelector('[data-element-id]')).toBeNull();
		expect(host.querySelector('[data-pptx-element]')).toBeNull();
		expect(host.querySelector('[data-pptx-viewport]')).toBeNull();
		expect(host.querySelector('[aria-roledescription="slide"]')).toBeNull();
		expect(stage(fixture)?.classList.contains('is-editable')).toBeFalsy();
	});

	it('follows edits to the slide model without a save', () => {
		const fixture = mount([shape('a')]);
		expect(drawn(fixture)).toHaveLength(1);
		fixture.componentInstance.slide.set(slide([shape('a'), shape('b', 300)]));
		fixture.detectChanges();
		expect(drawn(fixture)).toHaveLength(2);
	});

	it('follows slide metadata edits that keep the same elements', () => {
		const number = {
			...shape('number'),
			type: 'text',
			textSegments: [{ text: '#', style: {}, fieldType: 'slidenum' }],
		} as unknown as PptxElement;
		const fixture = mount([number]);
		const base = fixture.componentInstance.slide();
		expect(stage(fixture)?.textContent?.trim()).toBe('1');
		fixture.componentInstance.slide.set({ ...base, slideNumber: 7 });
		fixture.detectChanges();
		expect(stage(fixture)?.textContent?.trim()).toBe('7');
	});

	// The same badge, from the same shared descriptor, in every binding.
	it('draws the transition badge in the corner, at the slide scale', () => {
		const fixture = mount([shape('a')]);
		const badge = () =>
			box(fixture)?.querySelector<HTMLElement>(':scope > [data-pptx-thumbnail-transition]');
		const base = fixture.componentInstance.slide();
		expect(badge()).toBeNull();
		fixture.componentInstance.slide.set({ ...base, transition: { type: 'fade' } });
		fixture.detectChanges();
		expect(badge()?.textContent?.trim()).toBe('Fade');
		expect(badge()?.title).toBe('Transition: Fade');
		expect(badge()?.style.transform).toBe('scale(0.125)');
		// The drawn slide's corner, 120px from the box's left edge.
		expect(badge()?.style.right).toBe('calc(100% - 119.75px)');
		fixture.componentInstance.slide.set({ ...base, transition: { type: 'cut' } });
		fixture.detectChanges();
		expect(badge()).toBeNull();
	});
});
