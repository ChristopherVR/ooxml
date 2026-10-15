import { signal } from '@angular/core';
import { TestBed } from '@angular/core/testing';
import { BrowserTestingModule, platformBrowserTesting } from '@angular/platform-browser/testing';
import { TranslateService } from '@ngx-translate/core';
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { ReviewCommentsNavigationService } from './review-comments-navigation.service';
import { RibbonReviewSectionComponent } from './ribbon-review-section.component';

registerPptxWebControls();
beforeAll(() => TestBed.initTestEnvironment(BrowserTestingModule, platformBrowserTesting()));
afterEach(() => TestBed.resetTestingModule());
function mount() {
	TestBed.configureTestingModule({
		imports: [RibbonReviewSectionComponent],
		providers: [{ provide: TranslateService, useValue: { instant: (key: string) => key } }],
	});
	const fixture = TestBed.createComponent(RibbonReviewSectionComponent);
	fixture.componentRef.setInput('canEdit', true);
	fixture.detectChanges();
	const el = fixture.nativeElement as HTMLElement;
	const button = (id: string) =>
		el
			.querySelector(`[data-ribbon-control="review.${id}"]`)!
			.shadowRoot!.querySelector<HTMLButtonElement>('button')!;
	return { fixture, el, button };
}

describe('review shared command adapter', () => {
	it('emits the native outputs and leaves unsupported commands inert', () => {
		const { fixture, el, button } = mount();
		const spelling = vi.fn(),
			comments = vi.fn(),
			compare = vi.fn(),
			accessibility = vi.fn(),
			language = vi.fn();
		fixture.componentInstance.spellCheckChange.subscribe(spelling);
		fixture.componentInstance.comments.subscribe(comments);
		fixture.componentInstance.openCompare.subscribe(compare);
		fixture.componentInstance.a11y.subscribe(accessibility);
		fixture.componentInstance.language.subscribe(language);
		for (const id of [
			'proofing.spelling',
			'accessibility.check',
			'language.language',
			'compare.compare',
			'comments.newComment',
			'comments.showComments',
		]) {
			button(id).click();
		}
		expect(spelling).toHaveBeenCalledExactlyOnceWith(true);
		expect(accessibility).toHaveBeenCalledOnce();
		expect(language).toHaveBeenCalledOnce();
		expect(compare).toHaveBeenCalledOnce();
		expect(comments).toHaveBeenCalledTimes(2);
		expect(el.querySelectorAll('pptx-ui-ribbon-group')).toHaveLength(7);
		for (const id of [
			'proofing.thesaurus',
			'language.translate',
			'comments.delete',
			'protect.readOnly',
			'ink.hideInk',
		]) {
			expect(button(id).disabled).toBeTruthy();
		}
	});

	it('updates spelling and compare availability without replacing a focused command', () => {
		const { fixture, button } = mount();
		const focused = button('proofing.spelling');
		focused.focus();
		fixture.componentRef.setInput('spellCheckEnabled', true);
		fixture.componentRef.setInput('canEdit', false);
		fixture.detectChanges();
		expect(button('proofing.spelling')).toBe(focused);
		expect(focused.getAttribute('aria-pressed')).toBe('true');
		expect(button('compare.compare').disabled).toBeTruthy();
	});

	it('walks the deck comments and deletes the current one', async () => {
		const slides = signal([
			{ id: 's1', elements: [], comments: [{ id: 'a', text: 'a' }] },
			{ id: 's2', elements: [], comments: [{ id: 'b', text: 'b' }] },
		] as unknown as readonly PptxSlide[]);
		const applyReplacement = vi.fn((next: readonly PptxSlide[]) => slides.set(next));
		TestBed.configureTestingModule({
			imports: [RibbonReviewSectionComponent],
			providers: [
				{ provide: TranslateService, useValue: { instant: (key: string) => key } },
				{ provide: EditorStateService, useValue: { slides, applyReplacement } },
				ReviewCommentsNavigationService,
			],
		});
		const nav = TestBed.inject(ReviewCommentsNavigationService);
		const goTo = vi.fn();
		nav.attach(goTo);
		const fixture = TestBed.createComponent(RibbonReviewSectionComponent);
		fixture.componentRef.setInput('canEdit', true);
		fixture.detectChanges();
		const button = (id: string) =>
			(fixture.nativeElement as HTMLElement)
				.querySelector(`[data-ribbon-control="review.${id}"]`)!
				.shadowRoot!.querySelector<HTMLButtonElement>('button')!;
		expect(button('comments.delete').disabled).toBeTruthy();
		button('comments.next').click();
		fixture.detectChanges();
		button('comments.next').click();
		fixture.detectChanges();
		expect(goTo).toHaveBeenCalledExactlyOnceWith(1);
		expect(button('comments.delete').disabled).toBeFalsy();
		button('comments.delete').click();
		expect(applyReplacement).toHaveBeenCalledOnce();
		expect(slides()[1]?.comments).toStrictEqual([]);
	});
});
