/**
 * The shape panel's Quick Styles are the shared gallery (`pptx-ui-ribbon-gallery`, a compact
 * panel) with a caption in the current language, and a pick merges the preset over the shape's
 * own style.
 */
import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { SHAPE_QUICK_STYLES, shapeQuickStylePatch } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { translationsDe } from 'pptx-viewer-locales';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { registerPptxWebControls } from '../../../../../../src/ui/src/pptx/web-components';
import { ShapeAuthoringPanelComponent } from './shape-authoring-panel.component';

beforeAll(() => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	registerPptxWebControls();
});
afterEach(() => TestBed.resetTestingModule());

const shape = {
	id: 'box',
	type: 'shape',
	x: 0,
	y: 0,
	width: 200,
	height: 100,
	shapeType: 'rect',
	shapeStyle: { fillMode: 'solid', fillColor: '#FF0000', strokeWidth: 3, shadowBlur: 9 },
} as PptxElement;

function mount(element: PptxElement, language = 'en') {
	TestBed.configureTestingModule({
		imports: [ShapeAuthoringPanelComponent],
		providers: [provideTranslateService({ fallbackLang: 'en' })],
	});
	// The Vitest JIT build does not wire signal inputs, so declare them and hand each
	// field a plain signal.
	TestBed.overrideComponent(ShapeAuthoringPanelComponent, { add: { inputs: ['element'] } });
	const translate = TestBed.inject(TranslateService);
	translate.setTranslation('en', translationsEn);
	translate.setTranslation('de', translationsDe);
	translate.use(language);
	const fixture = TestBed.createComponent(ShapeAuthoringPanelComponent);
	fixture.componentRef.setInput('element', signal(element));
	fixture.detectChanges();
	const host = fixture.nativeElement as HTMLElement;
	const tiles = () => [...host.querySelectorAll<HTMLButtonElement>('[data-gallery-item]')];
	return { fixture, host, tiles, translate };
}

describe('shapeAuthoringPanelComponent Quick Styles', () => {
	it('draws every preset in order on the shared compact panel', () => {
		const { host, tiles } = mount(shape);
		const gallery = host.querySelector('pptx-ui-ribbon-gallery')!;
		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(tiles()).toHaveLength(18);
		expect(tiles().map((tile) => tile.getAttribute('aria-label'))).toStrictEqual(
			SHAPE_QUICK_STYLES.map((preset) => preset.name),
		);
		expect(tiles()[0]!.querySelector('svg')).not.toBeNull();
	});

	it('captions the gallery in the current language', () => {
		const { host } = mount(shape, 'de');
		expect(host.textContent).not.toMatch(/quick styles/iu);
		const caption = host.querySelector('pptx-ui-ribbon-gallery')!.previousElementSibling!;
		expect(caption.textContent?.trim()).toBe(translationsDe['pptx.shape.quickStyles']);
	});

	it("emits the preset merged over the shape's own style", () => {
		const { fixture, tiles } = mount(shape);
		const patches: Partial<PptxElement>[] = [];
		fixture.componentInstance.patch.subscribe((patch) => patches.push(patch));
		tiles()[0]!.click();
		tiles()[12]!.click();
		expect(patches).toStrictEqual([
			shapeQuickStylePatch(shape, SHAPE_QUICK_STYLES[0]!.style),
			shapeQuickStylePatch(shape, SHAPE_QUICK_STYLES[12]!.style),
		]);
		// The merge keeps what a preset does not set.
		expect((patches[0] as { shapeStyle: { shadowBlur?: number } }).shapeStyle.shadowBlur).toBe(9);
	});
});
