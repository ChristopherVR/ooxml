import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { FindReplaceState } from './editor-find-replace.svelte';

/**
 * FindReplaceState is a runes class (`.svelte.ts`); this suite is named
 * `.svelte.test.ts` so the module is compiled with the runes runtime.
 */

function textEl(id: string, text: string): PptxElement {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		text,
		textSegments: [{ text, style: {} }],
		textStyle: {},
	} as PptxElement;
}

function slide(elements: PptxElement[]): PptxSlide {
	return { id: 's', rId: 'rId-s', slideNumber: 1, elements };
}

describe('findReplaceState', () => {
	it('searches on toggle-open and reports match count', () => {
		let slides: PptxSlide[] = [slide([textEl('a', 'hello world'), textEl('b', 'hello again')])];
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: (next) => {
				slides = next;
			},
		});
		state.query = 'hello';
		state.toggle();
		expect(state.open).toBeTruthy();
		expect(state.matchCount).toBe(2);
		expect(state.hasResults).toBeTruthy();
	});

	it('navigates next/prev with wraparound', () => {
		let slides: PptxSlide[] = [slide([textEl('a', 'aa aa aa')])];
		const onNavigate = vi.fn();
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: (next) => {
				slides = next;
			},
			onNavigate,
		});
		state.query = 'aa';
		state.search();
		expect(state.results.length).toBeGreaterThan(1);
		state.next();
		expect(state.index).toBe(1);
		state.prev();
		state.prev();
		expect(state.index).toBe(state.results.length - 1);
	});

	it('replaceCurrent replaces one match and re-searches', () => {
		let slides: PptxSlide[] = [slide([textEl('a', 'hello world')])];
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: (next) => {
				slides = next;
			},
		});
		state.query = 'hello';
		state.replacement = 'hi';
		state.search();
		state.replaceCurrent();
		expect((slides[0].elements[0] as PptxElement & { text: string }).text).toBe('hi world');
		expect(state.matchCount).toBe(0);
	});

	it('replaceAll replaces every match', () => {
		let slides: PptxSlide[] = [slide([textEl('a', 'cat cat cat')])];
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: (next) => {
				slides = next;
			},
		});
		state.query = 'cat';
		state.replacement = 'dog';
		state.replaceAll();
		expect((slides[0].elements[0] as PptxElement & { text: string }).text).toBe('dog dog dog');
	});

	it('does nothing on replace when there are no results', () => {
		const commitSlides = vi.fn();
		const state = new FindReplaceState({ getSlides: () => [], commitSlides });
		state.replaceCurrent();
		expect(commitSlides).not.toHaveBeenCalled();
	});

	it('closing drops the matches; reopening searches the kept query from the first', () => {
		const slides: PptxSlide[] = [slide([textEl('a', 'foo foo foo')])];
		const state = new FindReplaceState({ getSlides: () => slides, commitSlides: () => {} });
		state.toggle();
		state.query = 'foo';
		state.search();
		state.next();
		state.close();
		expect(state.open).toBeFalsy();
		expect(state.matchCount).toBe(0);
		state.toggle();
		expect(state.query).toBe('foo');
		expect(state.matchCount).toBe(3);
		expect(state.index).toBe(0);
	});

	it('replace changes the current match and moves on to the next one', () => {
		let slides: PptxSlide[] = [
			{ ...slide([textEl('a', 'foo bar')]), id: 's1' },
			{ ...slide([textEl('b', 'foo foo')]), id: 's2' },
		];
		const onNavigate = vi.fn();
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: (next) => {
				slides = next;
			},
			onNavigate,
		});
		state.query = 'foo';
		state.replacement = 'qux';
		state.search();
		state.next();
		state.replaceCurrent();
		expect(slides.map((s) => (s.elements[0] as PptxElement & { text: string }).text)).toStrictEqual(
			['foo bar', 'qux foo'],
		);
		expect(state.matchCount).toBe(2);
		expect(state.index).toBe(1);
		expect(onNavigate).toHaveBeenLastCalledWith(1, 'b');
	});

	it('replace refuses a match the deck no longer has where the search found it', () => {
		let slides: PptxSlide[] = [slide([textEl('a', 'foo bar')])];
		const commitSlides = vi.fn((next: PptxSlide[]) => {
			slides = next;
		});
		const state = new FindReplaceState({ getSlides: () => slides, commitSlides });
		state.query = 'foo';
		state.replacement = 'qux';
		state.search();
		slides = [slide([textEl('a', 'xx foo bar')])];
		state.replaceCurrent();
		expect(commitSlides).not.toHaveBeenCalled();
		expect((slides[0].elements[0] as PptxElement & { text: string }).text).toBe('xx foo bar');
	});

	it('a speaker-notes match goes to its slide and selects no element', () => {
		const slides: PptxSlide[] = [slide([]), { ...slide([]), notes: 'speaker foo' }];
		const onNavigate = vi.fn();
		const state = new FindReplaceState({
			getSlides: () => slides,
			commitSlides: () => {},
			onNavigate,
		});
		state.includeNotes = true;
		state.query = 'foo';
		state.search();
		expect(onNavigate).toHaveBeenLastCalledWith(1, null);
	});
});
