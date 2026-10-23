import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { applyFindReplacements } from './find-replace';
import {
	DEFAULT_FIND_REPLACE_OPTIONS,
	findReplaceStatus,
	findReplaceTarget,
	INITIAL_FIND_REPLACE_SESSION,
	planFindReplace,
	reduceFindReplaceSession,
} from './find-replace-session';
import type { FindReplaceSession, FindReplaceSessionEvent } from './find-replace-session';

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

function slide(id: string, elements: PptxElement[], notes?: string): PptxSlide {
	return { id, rId: `rId-${id}`, slideNumber: 1, elements, ...(notes ? { notes } : {}) };
}

/** A chart whose title is searchable but must be edited at its source. */
function linkedChart(id: string, title: string): PptxElement {
	return {
		type: 'chart',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		chartData: { title, titleTextEditable: false, series: [], categories: [] },
	} as unknown as PptxElement;
}

/** Three matches of "foo" on two slides. */
const DECK = [slide('a', [text('a1', 'foo bar foo')]), slide('b', [text('b1', 'a foo')])];

const run = (from: FindReplaceSession, ...events: FindReplaceSessionEvent[]): FindReplaceSession =>
	events.reduce(reduceFindReplaceSession, from);
const searched = (slides: readonly PptxSlide[] = DECK, query = 'foo') =>
	run(INITIAL_FIND_REPLACE_SESSION, { type: 'search', slides, query });

/** Apply a plan and land the replaced slides, as a host does. */
function replaceAndLand(session: FindReplaceSession, scope: 'current' | 'all', slides = DECK) {
	const plan = planFindReplace(session, slides, scope);
	if (!plan) {
		throw new Error('expected a plan');
	}
	const replaced = applyFindReplacements(slides, plan.matches, plan.replacement).slides;
	return {
		slides: replaced,
		session: run(
			session,
			{ type: 'replaced', plan, slides },
			{ type: 'slidesChanged', slides: replaced },
		),
	};
}

const texts = (slides: readonly PptxSlide[]) =>
	slides.flatMap((s) => s.elements.map((e) => (e as { text?: string }).text));

describe('find-replace session defaults', () => {
	it("starts empty with React's scopes: match case, notes and chart data off, chart titles on", () => {
		expect(INITIAL_FIND_REPLACE_SESSION).toMatchObject({
			query: '',
			replacement: '',
			matches: [],
			current: -1,
			pendingRefind: null,
		});
		expect(DEFAULT_FIND_REPLACE_OPTIONS).toStrictEqual({
			matchCase: false,
			includeNotes: false,
			includeChartTitles: true,
			includeChartData: false,
		});
		expect(findReplaceStatus(INITIAL_FIND_REPLACE_SESSION)).toStrictEqual({ current: 0, total: 0 });
		expect(findReplaceTarget(INITIAL_FIND_REPLACE_SESSION)).toBeNull();
	});
});

describe('reduceFindReplaceSession: search and navigation', () => {
	it('finds every match and makes the first one current', () => {
		const session = searched();
		expect(session.matches).toHaveLength(3);
		expect(findReplaceStatus(session)).toStrictEqual({ current: 1, total: 3 });
		expect(findReplaceTarget(session)).toStrictEqual({ slideIndex: 0, elementId: 'a1' });
	});

	it('Next and Previous wrap around both ends', () => {
		const session = searched();
		const next = run(session, { type: 'step', direction: 1 });
		expect(findReplaceStatus(next)).toStrictEqual({ current: 2, total: 3 });
		const last = run(next, { type: 'step', direction: 1 });
		expect(findReplaceTarget(last)).toStrictEqual({ slideIndex: 1, elementId: 'b1' });
		expect(findReplaceStatus(run(last, { type: 'step', direction: 1 })).current).toBe(1);
		expect(findReplaceStatus(run(session, { type: 'step', direction: -1 })).current).toBe(3);
	});

	it('a step with no matches, or onto the only match, changes nothing', () => {
		expect(
			reduceFindReplaceSession(INITIAL_FIND_REPLACE_SESSION, { type: 'step', direction: 1 }),
		).toBe(INITIAL_FIND_REPLACE_SESSION);
		const single = searched(DECK, 'bar');
		expect(reduceFindReplaceSession(single, { type: 'step', direction: -1 })).toBe(single);
	});

	it('a scope change searches again with the other scopes kept, and an empty query finds nothing', () => {
		const deck = [slide('a', [text('a1', 'Foo')], 'foo in the notes')];
		const session = run(searched(deck), {
			type: 'search',
			slides: deck,
			options: { includeNotes: true },
		});
		expect(session.options).toStrictEqual({ ...DEFAULT_FIND_REPLACE_OPTIONS, includeNotes: true });
		expect(session.matches).toHaveLength(2);
		const caseSensitive = run(session, {
			type: 'search',
			slides: deck,
			options: { matchCase: true },
		});
		expect(caseSensitive.matches).toHaveLength(1);
		expect(caseSensitive.options.includeNotes).toBeTruthy();
		const cleared = run(caseSensitive, { type: 'search', slides: deck, query: '' });
		expect(cleared).toMatchObject({ query: '', matches: [], current: -1 });
		expect(findReplaceStatus(cleared)).toStrictEqual({ current: 0, total: 0 });
	});

	it('a speaker-notes match goes to its slide and selects no element', () => {
		const deck = [slide('a', [], 'speaker foo')];
		const session = run(INITIAL_FIND_REPLACE_SESSION, {
			type: 'search',
			slides: deck,
			query: 'foo',
			options: { includeNotes: true },
		});
		expect(findReplaceTarget(session)).toStrictEqual({ slideIndex: 0, elementId: null });
	});

	it('closing drops the matches and keeps the query, replacement and scopes', () => {
		const session = run(
			searched(),
			{ type: 'setReplacement', replacement: 'qux' },
			{ type: 'search', slides: DECK, options: { matchCase: true } },
		);
		const closed = run(session, { type: 'close' });
		expect(closed).toMatchObject({ matches: [], current: -1, pendingRefind: null });
		expect(closed).toMatchObject({ query: 'foo', replacement: 'qux' });
		expect(closed.options.matchCase).toBeTruthy();
		expect(reduceFindReplaceSession(closed, { type: 'close' })).toBe(closed);
	});

	it('setting the same replacement changes nothing', () => {
		const session = run(searched(), { type: 'setReplacement', replacement: 'x' });
		expect(reduceFindReplaceSession(session, { type: 'setReplacement', replacement: 'x' })).toBe(
			session,
		);
	});
});

describe('planFindReplace', () => {
	it('Replace plans the current match only', () => {
		const session = run(searched(), { type: 'step', direction: 1 });
		const plan = planFindReplace(session, DECK, 'current');
		expect(plan?.matches).toStrictEqual([session.matches[1]]);
		expect(plan?.scope).toBe('current');
	});

	it('Replace All plans every match', () => {
		expect(planFindReplace(searched(), DECK, 'all')?.matches).toHaveLength(3);
	});

	it('refuses when there is nothing to replace', () => {
		expect(planFindReplace(INITIAL_FIND_REPLACE_SESSION, DECK, 'current')).toBeNull();
		expect(planFindReplace(searched(DECK, 'zzz'), DECK, 'all')).toBeNull();
	});

	it('refuses Replace All when any match is read-only, and Replace on a read-only match', () => {
		const deck = [slide('a', [text('a1', 'report'), linkedChart('c1', 'Annual report')])];
		const session = searched(deck, 'report');
		expect(session.matches.map((match) => Boolean(match.readOnly))).toStrictEqual([false, true]);
		expect(planFindReplace(session, deck, 'all')).toBeNull();
		expect(planFindReplace(session, deck, 'current')?.matches).toHaveLength(1);
		const onChart = run(session, { type: 'step', direction: 1 });
		expect(planFindReplace(onChart, deck, 'current')).toBeNull();
	});

	it('refuses Replace when the deck changed and the current match is no longer there', () => {
		const session = searched();
		const edited = [slide('a', [text('a1', 'xxfoo bar foo')]), DECK[1]!];
		expect(planFindReplace(session, edited, 'current')).toBeNull();
		// Replace All searches the deck as it is now.
		expect(planFindReplace(session, edited, 'all')?.matches).toHaveLength(3);
	});
});

describe('reduceFindReplaceSession: searching again after a replace', () => {
	it('Replace changes only the current match and moves on to the next one', () => {
		const session = run(
			searched(),
			{ type: 'setReplacement', replacement: 'qux' },
			{
				type: 'step',
				direction: 1,
			},
		);
		const after = replaceAndLand(session, 'current');
		expect(texts(after.slides)).toStrictEqual(['foo bar qux', 'a foo']);
		expect(findReplaceStatus(after.session)).toStrictEqual({ current: 2, total: 2 });
		expect(findReplaceTarget(after.session)).toStrictEqual({ slideIndex: 1, elementId: 'b1' });
		expect(after.session.pendingRefind).toBeNull();
	});

	it('Replace on the last match wraps to the first', () => {
		const session = run(searched(), { type: 'step', direction: -1 });
		const after = replaceAndLand(session, 'current');
		expect(texts(after.slides)).toStrictEqual(['foo bar foo', 'a ']);
		expect(findReplaceStatus(after.session)).toStrictEqual({ current: 1, total: 2 });
	});

	it('Replace steps past a match the replacement text itself created', () => {
		const session = run(searched(), { type: 'setReplacement', replacement: 'food' });
		const after = replaceAndLand(session, 'current');
		expect(texts(after.slides)).toStrictEqual(['food bar foo', 'a foo']);
		// "food" still holds "foo": the cursor skips it to the old second match.
		expect(findReplaceStatus(after.session)).toStrictEqual({ current: 2, total: 3 });
		expect(after.session.matches[1]?.startOffset).toBe(9);
	});

	it('Replace All counts again from the first match', () => {
		const session = run(searched(), { type: 'step', direction: 1 });
		const gone = replaceAndLand(
			run(session, { type: 'setReplacement', replacement: 'qux' }),
			'all',
		);
		expect(texts(gone.slides)).toStrictEqual(['qux bar qux', 'a qux']);
		expect(findReplaceStatus(gone.session)).toStrictEqual({ current: 0, total: 0 });
		const kept = replaceAndLand(
			run(session, { type: 'setReplacement', replacement: 'food' }),
			'all',
		);
		expect(findReplaceStatus(kept.session)).toStrictEqual({ current: 1, total: 3 });
	});

	it('Replace after an ordinary edit added a match moves on from the deck as it is now', () => {
		const deck = [
			slide('a', [text('a1', 'foo')]),
			slide('b', [text('b1', 'foo')]),
			slide('c', [text('c1', 'bar')]),
		];
		const session = run(searched(deck), { type: 'setReplacement', replacement: 'qux' });
		expect(session.matches).toHaveLength(2);
		// An ordinary edit adds a match on slide c; the session keeps its two matches.
		const edited = [deck[0]!, deck[1]!, slide('c', [text('c1', 'foo')])];
		expect(reduceFindReplaceSession(session, { type: 'slidesChanged', slides: edited })).toBe(
			session,
		);
		const after = replaceAndLand(session, 'current', edited);
		expect(texts(after.slides)).toStrictEqual(['qux', 'foo', 'foo']);
		// The match after the replaced one is b's, not c's.
		expect(findReplaceTarget(after.session)).toStrictEqual({ slideIndex: 1, elementId: 'b1' });
		expect(findReplaceStatus(after.session)).toStrictEqual({ current: 1, total: 2 });
	});

	it('Replace after an ordinary edit steps past a match the replacement itself created', () => {
		const deck = [
			slide('a', [text('a1', 'foo')]),
			slide('b', [text('b1', 'foo')]),
			slide('c', [text('c1', 'bar')]),
		];
		const session = run(searched(deck), { type: 'setReplacement', replacement: 'food' });
		const edited = [deck[0]!, deck[1]!, slide('c', [text('c1', 'foo')])];
		const after = replaceAndLand(session, 'current', edited);
		expect(texts(after.slides)).toStrictEqual(['food', 'foo', 'foo']);
		expect(findReplaceTarget(after.session)).toStrictEqual({ slideIndex: 1, elementId: 'b1' });
		expect(findReplaceStatus(after.session)).toStrictEqual({ current: 2, total: 3 });
	});

	it('waits while the host still holds the replaced slides, then searches once they land', () => {
		const session = searched();
		const plan = planFindReplace(session, DECK, 'current')!;
		const pending = run(session, { type: 'replaced', plan, slides: DECK });
		expect(pending.pendingRefind).toMatchObject({ scope: 'current', index: 0, count: 3 });
		expect(reduceFindReplaceSession(pending, { type: 'slidesChanged', slides: DECK })).toBe(
			pending,
		);
		const landed = applyFindReplacements(DECK, plan.matches, '').slides;
		expect(run(pending, { type: 'slidesChanged', slides: landed }).matches).toHaveLength(2);
	});

	it('an ordinary edit with no replace waiting does not search again', () => {
		const session = run(searched(), { type: 'step', direction: 1 });
		const edited = [slide('a', [text('a1', 'foo')]), DECK[1]!];
		expect(reduceFindReplaceSession(session, { type: 'slidesChanged', slides: edited })).toBe(
			session,
		);
	});

	it('a new search or a close drops a replace still waiting for its slides', () => {
		const session = searched();
		const plan = planFindReplace(session, DECK, 'all')!;
		const pending = run(session, { type: 'replaced', plan, slides: DECK });
		expect(run(pending, { type: 'close' }).pendingRefind).toBeNull();
		expect(run(pending, { type: 'search', slides: DECK }).pendingRefind).toBeNull();
	});
});
