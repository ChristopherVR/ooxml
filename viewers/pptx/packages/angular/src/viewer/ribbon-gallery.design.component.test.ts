/**
 * Design > Variants > Background Styles in `<pptx-ribbon-gallery>`: a hovered tile is shown over
 * the slides without an undo step (and holds collaboration), leaving puts the original back, a
 * pick commits one undo step, and the footer command is emitted for the Design section.
 */
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { provideTranslateService, TranslateService } from '@ngx-translate/core';
import type { PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';

import { translationsEn } from '../../../../../../src/ui/src/pptx/i18n/translations-en';
import { resolveViewerComponentResources } from './component-resources.test-support';
import { EditorStateService } from './editor-state.service';
import { LoadContentService } from './load-content.service';
import { RibbonGalleryComponent } from './ribbon-gallery.component';
import { ViewerDeckStyleService } from './viewer-deck-style.service';

beforeAll(async () => {
	TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting());
	await resolveViewerComponentResources();
});
beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	vi.useRealTimers();
	TestBed.resetTestingModule();
});

const handler = {
	resolveStyleMatrixReferences: () => ({ shapeStyle: {} }),
	getThemeVariants: () => [],
	applyBackgroundStyleToSlides: (slides: PptxSlide[], templates: unknown, index: number) => ({
		slides: slides.map((slide) => ({ ...slide, backgroundColor: `#00000${index}` })),
		templateElementsBySlideId: templates,
	}),
	commitMasterBackgroundStyle: vi.fn(async () => true),
};

function mount() {
	TestBed.configureTestingModule({
		imports: [RibbonGalleryComponent],
		providers: [
			provideTranslateService({ fallbackLang: 'en' }),
			EditorStateService,
			ViewerDeckStyleService,
			LoadContentService,
		],
	});
	TestBed.inject(TranslateService).setTranslation('en', translationsEn);
	TestBed.inject(TranslateService).use('en');
	const loader = TestBed.inject(LoadContentService);
	vi.spyOn(loader, 'getHandler').mockReturnValue(handler as never);
	loader.theme.set({ name: 'Office' } as PptxTheme);
	const editor = TestBed.inject(EditorStateService);
	editor.setSlides([
		{
			id: 's1',
			slideNumber: 1,
			elements: [],
			backgroundColor: '#FFFFFF',
		} as unknown as PptxSlide,
	]);
	const fixture = TestBed.createComponent(RibbonGalleryComponent);
	fixture.componentRef.setInput('gallery', 'themeBackgrounds');
	fixture.componentRef.setInput('mode', 'dropdown');
	fixture.componentRef.setInput('canEdit', true);
	fixture.componentRef.setInput('slideIndex', 0);
	fixture.detectChanges();
	const root = fixture.nativeElement as HTMLElement;
	document.body.append(root);
	root.querySelector<HTMLButtonElement>('[data-ribbon-gallery]')!.click();
	fixture.detectChanges();
	return { root, editor, fixture };
}

describe('Design > Background Styles (Angular)', () => {
	it('shows a hovered style, holds collaboration and restores the original on leave', () => {
		const { root, editor, fixture } = mount();
		const original = editor.slides();
		const tile = root.querySelector<HTMLElement>('[data-gallery-item="3"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		fixture.detectChanges();
		expect(editor.slides()[0]!.backgroundColor).toBe('#000003');
		expect(editor.deckStyleHeld()).toBe(true);
		expect(editor.dirty()).toBe(false);
		tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(100);
		expect(editor.slides()).toBe(original);
		expect(editor.deckStyleHeld()).toBe(false);
		expect(editor.dirty()).toBe(false);
		root.remove();
	});

	it('keeps a hover preview out of what a save or recovery snapshot writes', () => {
		const { root, editor, fixture } = mount();
		const service = TestBed.inject(ViewerDeckStyleService);
		const original = editor.slides();
		const tile = root.querySelector<HTMLElement>('[data-gallery-item="3"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		fixture.detectChanges();
		expect(editor.slides()[0]!.backgroundColor).toBe('#000003');
		expect(service.savedDeck().slides).toBe(original);
		tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(100);
		expect(service.savedDeck().slides).toBe(original);
		root.remove();
	});

	it('commits a pick as one undo step from the real slides', async () => {
		const { root, editor, fixture } = mount();
		const tile = root.querySelector<HTMLElement>('[data-gallery-item="6"]')!;
		tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
		vi.advanceTimersByTime(150);
		tile.click();
		await vi.advanceTimersByTimeAsync(0);
		fixture.detectChanges();
		expect(editor.slides()[0]!.backgroundColor).toBe('#000006');
		expect(editor.dirty()).toBe(true);
		expect(editor.deckStyleHeld()).toBe(false);
		editor.undo();
		expect(editor.slides()[0]!.backgroundColor).toBe('#FFFFFF');
		root.remove();
	});

	it('emits the Format Background command from the footer', () => {
		const { root, fixture } = mount();
		const commands: string[] = [];
		fixture.componentInstance.command.subscribe((command) => commands.push(command));
		root.querySelector<HTMLButtonElement>('[data-gallery-footer="formatBackground"]')!.click();
		expect(commands).toEqual(['formatBackground']);
		root.remove();
	});
});
