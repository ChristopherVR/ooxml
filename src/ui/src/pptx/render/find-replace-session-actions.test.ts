import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it, vi } from 'vitest';

import { INITIAL_FIND_REPLACE_SESSION, findReplaceStatus } from './find-replace-session';
import type { FindReplaceSession, FindReplaceTarget } from './find-replace-session';
import { findReplaceSessionActions } from './find-replace-session-actions';
import type { FindReplaceSessionHost } from './find-replace-session-actions';

function text(id: string, value: string): PptxElement {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 20,
		text: value,
		textSegments: [{ text: value, style: {} }],
	} as PptxElement;
}

const deck = (): PptxSlide[] => [
	{ id: 'a', rId: 'rId-a', slideNumber: 1, elements: [text('a1', 'foo bar foo')] },
	{ id: 'b', rId: 'rId-b', slideNumber: 2, elements: [text('b1', 'a foo')] },
];

/**
 * A host whose slides land at once (`deferred: false`) or only when the test calls `land()`, the
 * way React's state does.
 */
function harness(deferred = false) {
	let session: FindReplaceSession = INITIAL_FIND_REPLACE_SESSION;
	let slides: readonly PptxSlide[] = deck();
	let queued: readonly PptxSlide[] | null = null;
	const navigate = vi.fn<(target: FindReplaceTarget) => void>();
	const commit = vi.fn<FindReplaceSessionHost['commit']>((next) => {
		if (deferred) {
			queued = next;
		} else {
			slides = next;
		}
	});
	const actions = findReplaceSessionActions(
		{ get: () => session, set: (next) => (session = next) },
		{ slides: () => slides, commit, navigate },
	);
	return {
		actions,
		navigate,
		commit,
		session: () => session,
		slides: () => slides,
		edit: (next: readonly PptxSlide[]) => (slides = next),
		land: () => {
			slides = queued ?? slides;
			queued = null;
			actions.slidesChanged();
		},
	};
}

const texts = (slides: readonly PptxSlide[]) =>
	slides.map((s) => (s.elements[0] as { text?: string }).text);

describe('findReplaceSessionActions', () => {
	it('a search shows the first match; Next and Previous show the one they land on', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		expect(h.navigate).toHaveBeenLastCalledWith({ slideIndex: 0, elementId: 'a1' });
		h.actions.next();
		h.actions.next();
		expect(h.navigate).toHaveBeenLastCalledWith({ slideIndex: 1, elementId: 'b1' });
		expect(findReplaceStatus(h.session())).toStrictEqual({ current: 3, total: 3 });
		h.actions.next();
		expect(findReplaceStatus(h.session()).current).toBe(1);
		h.actions.previous();
		expect(h.navigate).toHaveBeenLastCalledWith({ slideIndex: 1, elementId: 'b1' });
	});

	it('Next on the only match shows it again', () => {
		const h = harness();
		h.actions.search({ query: 'bar' });
		h.navigate.mockClear();
		h.actions.next();
		expect(h.navigate).toHaveBeenCalledExactlyOnceWith({ slideIndex: 0, elementId: 'a1' });
	});

	it('Replace commits the current match only and searches again at once on a synchronous host', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.actions.setReplacement('qux');
		h.actions.next();
		expect(h.actions.replaceCurrent()).toBe(1);
		expect(h.commit).toHaveBeenCalledExactlyOnceWith(expect.any(Array), 'current');
		expect(texts(h.slides())).toStrictEqual(['foo bar qux', 'a foo']);
		expect(findReplaceStatus(h.session())).toStrictEqual({ current: 2, total: 2 });
		expect(h.navigate).toHaveBeenLastCalledWith({ slideIndex: 1, elementId: 'b1' });
	});

	it('Replace All commits once with the scope, then finds no match left', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.actions.setReplacement('qux');
		expect(h.actions.replaceAll()).toBe(3);
		expect(h.commit).toHaveBeenCalledExactlyOnceWith(expect.any(Array), 'all');
		expect(texts(h.slides())).toStrictEqual(['qux bar qux', 'a qux']);
		expect(findReplaceStatus(h.session())).toStrictEqual({ current: 0, total: 0 });
	});

	it('a deferred host is searched again when its replaced slides land, not before', () => {
		const h = harness(true);
		h.actions.search({ query: 'foo' });
		h.actions.replaceAll();
		expect(h.session().matches).toHaveLength(3);
		expect(h.session().pendingRefind).not.toBeNull();
		h.actions.slidesChanged();
		expect(h.session().matches).toHaveLength(3);
		h.land();
		expect(h.session().matches).toHaveLength(0);
		expect(h.session().pendingRefind).toBeNull();
	});

	it('an ordinary edit neither searches again nor navigates', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.actions.next();
		h.navigate.mockClear();
		h.edit([{ ...deck()[0]!, elements: [text('a1', 'foo')] }, deck()[1]!]);
		h.actions.slidesChanged();
		expect(h.session().matches).toHaveLength(3);
		expect(findReplaceStatus(h.session()).current).toBe(2);
		expect(h.navigate).not.toHaveBeenCalled();
	});

	it('two Replaces after an ordinary edit added a match change the next two matches in order', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.actions.setReplacement('qux');
		// An ordinary edit adds a match on the second slide; the session keeps its three.
		h.edit([deck()[0]!, { ...deck()[1]!, elements: [text('b1', 'a foo foo')] }]);
		h.actions.slidesChanged();
		expect(h.actions.replaceCurrent()).toBe(1);
		expect(texts(h.slides())).toStrictEqual(['qux bar foo', 'a foo foo']);
		expect(findReplaceStatus(h.session())).toStrictEqual({ current: 1, total: 3 });
		expect(h.actions.replaceCurrent()).toBe(1);
		expect(texts(h.slides())).toStrictEqual(['qux bar qux', 'a foo foo']);
		expect(h.navigate).toHaveBeenLastCalledWith({ slideIndex: 1, elementId: 'b1' });
	});

	it('a refused replace commits nothing and leaves no search waiting', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.edit([{ ...deck()[0]!, elements: [text('a1', 'xx foo')] }, deck()[1]!]);
		expect(h.actions.replaceCurrent()).toBe(0);
		expect(h.commit).not.toHaveBeenCalled();
		expect(h.session().pendingRefind).toBeNull();
	});

	it('closing drops the matches without navigating; a later search starts from the first', () => {
		const h = harness();
		h.actions.search({ query: 'foo' });
		h.actions.next();
		h.navigate.mockClear();
		h.actions.close();
		expect(h.session()).toMatchObject({ query: 'foo', matches: [], current: -1 });
		expect(h.navigate).not.toHaveBeenCalled();
		h.actions.search();
		expect(findReplaceStatus(h.session())).toStrictEqual({ current: 1, total: 3 });
	});
});
