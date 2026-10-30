import { describe, expect, it } from 'vitest';
import type { PptxCustomShow, PptxElement, PptxSlide } from 'ooxml-core/pptx';

import {
	addSlidesToCustomShow,
	clearCustomShowActions,
	copyCustomShow,
	createCustomShow,
	customShowSlides,
	deleteCustomShow,
	moveCustomShowSlide,
	nextCustomShowId,
	pruneCustomShowSlides,
	removeSlidesFromCustomShow,
	renameCustomShow,
	setCustomShowSlides,
} from './custom-show-operations';

const show = (id: string, name: string, slideRIds: string[]): PptxCustomShow => ({
	id,
	name,
	slideRIds,
});

const slide = (rId: string, elements: unknown[] = []): PptxSlide =>
	({ rId, elements }) as unknown as PptxSlide;

describe('custom show operations', () => {
	it('allocates numeric ids above the largest in use', () => {
		expect(nextCustomShowId([])).toBe('0');
		expect(nextCustomShowId([show('0', 'a', []), show('4', 'b', []), show('x', 'c', [])])).toBe(
			'5',
		);
	});

	it('creates shows with unique names and never mutates the input', () => {
		const base = [show('0', 'Intro', ['rId2'])];
		const created = createCustomShow(base, 'intro', ['rId3']);
		expect(created.show).toEqual({ id: '1', name: 'intro (2)', slideRIds: ['rId3'] });
		expect(created.shows).toHaveLength(2);
		expect(base).toHaveLength(1);
	});

	it('renames, replaces, inserts, removes and reorders slides', () => {
		let shows = createCustomShow([], 'A', ['r1', 'r2', 'r3']).shows;
		shows = renameCustomShow(shows, '0', 'B').shows;
		expect(shows[0]?.name).toBe('B');
		shows = addSlidesToCustomShow(shows, '0', ['r9'], 1).shows;
		expect(shows[0]?.slideRIds).toEqual(['r1', 'r9', 'r2', 'r3']);
		shows = moveCustomShowSlide(shows, '0', 3, 0).shows;
		expect(shows[0]?.slideRIds).toEqual(['r3', 'r1', 'r9', 'r2']);
		shows = removeSlidesFromCustomShow(shows, '0', [0, 2]).shows;
		expect(shows[0]?.slideRIds).toEqual(['r1', 'r2']);
		shows = setCustomShowSlides(shows, '0', ['r2', 'r2']).shows;
		expect(shows[0]?.slideRIds).toEqual(['r2', 'r2']);
		expect(renameCustomShow(shows, 'nope', 'x').show).toBeUndefined();
	});

	it('copies a show under a new id without its raw XML', () => {
		const source = { ...show('2', 'Deck', ['r1']), rawXml: { '@_id': '2' } };
		const copied = copyCustomShow([source], '2');
		expect(copied.show).toEqual({ id: '3', name: 'Copy of Deck', slideRIds: ['r1'] });
		expect(copied.show?.rawXml).toBeUndefined();
	});

	it('deletes a show and prunes slide references to removed slides', () => {
		const shows = [show('0', 'A', ['r1', 'r2']), show('1', 'B', ['r2'])];
		expect(deleteCustomShow(shows, '0').shows.map((s) => s.id)).toEqual(['1']);
		const pruned = pruneCustomShowSlides(shows, [slide('r1')]);
		expect(pruned[0]?.slideRIds).toEqual(['r1']);
		expect(pruned[1]?.slideRIds).toEqual([]);
		expect(customShowSlides(shows[0] as PptxCustomShow, [slide('r2'), slide('r1')])).toHaveLength(
			2,
		);
	});

	it('clears actions and text links that jump into a removed show only', () => {
		const target = {
			id: 'a',
			type: 'shape',
			actionClick: { action: 'ppaction://customshow?id=3&return=true' },
			actionHover: { action: 'ppaction://customshow?id=4' },
		};
		const group = {
			id: 'g',
			type: 'group',
			children: [
				{
					id: 't',
					type: 'text',
					textSegments: [
						{ text: 'x', style: { hyperlink: 'h', hyperlinkAction: 'ppaction://customshow?id=3' } },
						{ text: 'y', style: { hyperlinkAction: 'ppaction://customshow?id=4' } },
					],
				},
			],
		};
		const slides = [slide('r1', [target, group]), slide('r2', [{ id: 'z', type: 'shape' }])];
		const { slides: out, cleared } = clearCustomShowActions(slides, '3');
		expect(cleared).toBe(2);
		const [first, second] = out as [PptxSlide, PptxSlide];
		const el = first.elements[0] as PptxElement;
		expect(el.actionClick).toBeUndefined();
		expect(el.actionHover).toBeDefined();
		const seg = (
			(first.elements[1] as unknown as { children: PptxElement[] }).children[0] as PptxElement
		).textSegments;
		expect(seg?.[0]?.style.hyperlinkAction).toBeUndefined();
		expect(seg?.[0]?.style.hyperlink).toBeUndefined();
		expect(seg?.[1]?.style.hyperlinkAction).toBeDefined();
		expect(second).toBe(slides[1]);
		const original = (slides[0] as PptxSlide).elements[0] as PptxElement;
		expect(original.actionClick).toBeDefined();
	});
});
