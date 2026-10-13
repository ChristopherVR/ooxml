/**
 * Picture Format > Size > Crop, rendered: the contextual section draws the
 * shared Size strip for Picture Format, and its Crop split button runs the
 * same crop commands as Home > Arrange (toggle, an aspect ratio in one undo
 * step, hidden with the host's `crop` button).
 */
import { signal } from '@angular/core';
import type { Type } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { EditorStateService } from './editor-state.service';
import { PictureCropService } from './picture-crop.service';
import { RecentColorsService } from './recent-colors.service';
import { RibbonContextualSectionComponent } from './ribbon-contextual-section.component';
import { RibbonPictureSizeComponent } from './ribbon-picture-size.component';

beforeAll(async () => {
	registerPptxWebControls();
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
afterEach(() => {
	TestBed.resetTestingModule();
});

const PICTURE = {
	type: 'picture',
	id: 'pic',
	x: 300,
	y: 0,
	width: 200,
	height: 100,
	imageData: 'data:image/png;base64,AAAA',
} as PptxElement;

interface Mounted {
	root: HTMLElement;
	editor: EditorStateService;
	crop: PictureCropService;
	detect: () => void;
}

/** Render `component` with signal inputs (this JIT runner has no signal-input transform). */
function mount<T>(component: Type<T>, inputs: Record<string, unknown>): Mounted {
	TestBed.configureTestingModule({
		imports: [component],
		providers: [
			provideTranslateService({ fallbackLang: 'en' }),
			EditorStateService,
			PictureCropService,
			{ provide: RecentColorsService, useValue: { recent: () => [], push: () => undefined } },
		],
	});
	TestBed.overrideComponent(component, { add: { inputs: Object.keys(inputs) } });
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
	const editor = TestBed.inject(EditorStateService);
	editor.setSlides([{ id: 's1', slideNumber: 1, elements: [PICTURE] } as PptxSlide]);
	editor.select(['pic']);
	const fixture = TestBed.createComponent(component);
	for (const [name, value] of Object.entries(inputs)) {
		fixture.componentRef.setInput(name, signal(value));
	}
	fixture.detectChanges();
	return {
		root: fixture.nativeElement as HTMLElement,
		editor,
		crop: TestBed.inject(PictureCropService),
		detect: () => fixture.detectChanges(),
	};
}

const mountSize = (extra: Record<string, unknown> = {}) =>
	mount(RibbonPictureSizeComponent, {
		slideIndex: 0,
		selectedElement: PICTURE,
		canEdit: true,
		hiddenActions: [],
		...extra,
	});

function q(root: HTMLElement, hook: string): HTMLButtonElement | null {
	return root.querySelector<HTMLButtonElement>(`[data-pptx-ribbon-control="${hook}"]`);
}

describe('picture Format > Size', () => {
	it('is a group of the Picture Format contextual section', () => {
		const { root } = mount(RibbonContextualSectionComponent, {
			tab: 'pictureFormat',
			selectedElement: PICTURE,
			slideIndex: 0,
			canEdit: true,
		});
		const group = root.querySelector('[data-ribbon-group="pictureFormat.size"]');
		expect(group?.querySelector('pptx-ribbon-picture-size')).not.toBeNull();
	});

	it('toggles crop mode from its Crop button', () => {
		const { root, crop, detect } = mountSize();
		const toggle = q(root, 'crop');
		expect(toggle?.closest('[data-ribbon-control="pictureFormat.size.crop"]')).not.toBeNull();
		expect(toggle?.disabled).toBeFalsy();
		toggle?.click();
		detect();
		expect(crop.isCropping('pic')).toBeTruthy();
		expect(toggle?.getAttribute('aria-pressed')).toBe('true');
		toggle?.click();
		expect(crop.state()).toBeNull();
	});

	it('crops to a 1:1 aspect ratio as one undo step', () => {
		const { root, editor } = mountSize();
		q(root, 'crop-menu')?.click();
		root.querySelector<HTMLButtonElement>('[data-pptx-crop-aspect="1:1"]')?.click();
		const pic = editor.slides()[0].elements[0];
		expect(pic.width).toBeCloseTo(pic.height);
		editor.undo();
		expect(editor.slides()[0].elements[0].width).toBe(200);
	});

	it('hides its Crop when the host hides crop', () => {
		const { root } = mountSize({ hiddenActions: ['crop'] });
		expect(q(root, 'crop')?.hidden).toBeTruthy();
		expect(q(root, 'crop-menu')?.hidden).toBeTruthy();
	});
});
