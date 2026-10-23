// oxlint-disable react-hooks/rules-of-hooks
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import type { Ref } from 'vue';
import { nextTick, ref } from 'vue';

import { useFindReplace } from './useFindReplace';
import { useSlideChangeSelection } from './useSlideChangeSelection';

/** A `showMatch` that only moves `activeSlideIndex`, for the tests that check navigation alone. */
const showOn =
	(activeSlideIndex: Ref<number>) =>
	(slideIndex: number): void => {
		activeSlideIndex.value = slideIndex;
	};

function textEl(id: string, text: string): PptxElement {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text,
		textSegments: [{ text, style: {} }],
	} as PptxElement;
}

function slide(id: string, elements: PptxElement[]): PptxSlide {
	return { id, elements } as PptxSlide;
}

/** Look up a text element by id and narrow it, for asserting on `.text` / `.textSegments`. */
function findTextElement(
	slides: PptxSlide[],
	slideIndex: number,
	elementId: string,
): Extract<PptxElement, { type: 'text' }> {
	return slides[slideIndex].elements?.find((e) => e.id === elementId) as Extract<
		PptxElement,
		{ type: 'text' }
	>;
}

function makeSlides(): PptxSlide[] {
	return [
		slide('s1', [textEl('a', 'Hello world'), textEl('b', 'nothing here')]),
		slide('s2', [
			textEl('c', 'world peace'),
			{ type: 'image', id: 'img', x: 0, y: 0, width: 1, height: 1 } as PptxElement,
		]),
	];
}

describe('useFindReplace', () => {
	it('exposes a reactive match count over the live slides, at per-occurrence precision', () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });

		fr.query.value = 'world';
		expect(fr.matchCount.value).toBe(2);
	});

	it('counts every occurrence within a single element, not one match per element', () => {
		const slides = ref([slide('s1', [textEl('a', 'world world world')])]),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });

		fr.query.value = 'world';
		// Per-occurrence: 3 matches inside one element, not 1.
		expect(fr.matchCount.value).toBe(3);
	});

	it('next() cycles matches and moves the active slide', () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });
		fr.query.value = 'world';

		fr.next();
		expect(fr.currentMatch.value).toBe(1);
		expect(activeSlideIndex.value).toBe(1);
		fr.next();
		expect(fr.currentMatch.value).toBe(0);
		expect(activeSlideIndex.value).toBe(0);
	});

	it('prev() cycles backwards', () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });
		fr.query.value = 'world';

		fr.prev();
		expect(fr.currentMatch.value).toBe(1);
		expect(activeSlideIndex.value).toBe(1);
	});

	it('respects case sensitivity', () => {
		const slides = ref([slide('s1', [textEl('a', 'World'), textEl('b', 'world')])]),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });

		fr.matchCase.value = true;
		fr.query.value = 'world';
		expect(fr.matchCount.value).toBe(1);

		fr.matchCase.value = false;
		expect(fr.matchCount.value).toBe(2);
	});

	it('replaceAll rewrites text + segments and snapshots history', () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			pushHistory = vi.fn(),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory });

		fr.query.value = 'world';
		fr.replacement.value = 'earth';
		fr.replaceAll();

		expect(pushHistory).toHaveBeenCalledOnce();
		expect(findTextElement(slides.value, 0, 'a').text).toBe('Hello earth');
		expect(findTextElement(slides.value, 0, 'a').textSegments?.[0].text).toBe('Hello earth');
		expect(findTextElement(slides.value, 1, 'c').text).toBe('earth peace');
		// Query no longer matches after replacement.
		expect(fr.matchCount.value).toBe(0);
	});

	it('replaceAll replaces every occurrence within a single element, not just the first', () => {
		const slides = ref([slide('s1', [textEl('a', 'world world world')])]),
			activeSlideIndex = ref(0),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory: () => {} });

		fr.query.value = 'world';
		fr.replacement.value = 'earth';
		fr.replaceAll();

		expect(findTextElement(slides.value, 0, 'a').text).toBe('earth earth earth');
	});

	it('replaceCurrent only rewrites the focused occurrence', () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			pushHistory = vi.fn(),
			fr = useFindReplace({ slides, showMatch: showOn(activeSlideIndex), pushHistory });

		fr.query.value = 'world';
		fr.replacement.value = 'earth';
		fr.replaceCurrent();

		expect(pushHistory).toHaveBeenCalledOnce();
		expect(findTextElement(slides.value, 0, 'a').text).toBe('Hello earth');
		// Second match untouched, so one match remains.
		expect(findTextElement(slides.value, 1, 'c').text).toBe('world peace');
		expect(fr.matchCount.value).toBe(1);
	});

	it('selects the matched element, also across a slide change, which still clears otherwise', async () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			selected = ref<string[]>(['b']),
			// The viewer's own rule: a slide change drops the selection unless Find made it for that slide.
			slideSelection = useSlideChangeSelection({
				activeSlideIndex,
				goTo: (index) => {
					activeSlideIndex.value = index;
				},
				selectedElementIds: selected,
			}),
			fr = useFindReplace({
				slides,
				showMatch: slideSelection.showSlideSelection,
				pushHistory: () => {},
			});
		fr.query.value = 'world';
		await nextTick();
		expect(selected.value).toStrictEqual(['a']);
		fr.next();
		await nextTick();
		expect(activeSlideIndex.value).toBe(1);
		expect(selected.value).toStrictEqual(['c']);
		activeSlideIndex.value = 0;
		await nextTick();
		expect(selected.value).toStrictEqual([]);
	});

	it('closing the bar drops the matches; reopening searches again from the first', async () => {
		const slides = ref(makeSlides()),
			activeSlideIndex = ref(0),
			open = ref(true),
			fr = useFindReplace({
				slides,
				showMatch: showOn(activeSlideIndex),
				pushHistory: () => {},
				open,
			});
		fr.query.value = 'world';
		fr.next();
		expect(fr.currentMatch.value).toBe(1);
		open.value = false;
		await nextTick();
		expect(fr.matchCount.value).toBe(0);
		open.value = true;
		await nextTick();
		expect(fr.matchCount.value).toBe(2);
		expect(fr.currentMatch.value).toBe(0);
	});
});
