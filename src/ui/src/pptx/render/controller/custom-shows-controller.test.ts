import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createCustomShowsController } from './custom-shows-controller';

const slide = (rId: string, elements: unknown[] = []): PptxSlide =>
	({ rId, elements }) as unknown as PptxSlide;
const slides = [slide('r1'), slide('r2'), slide('r3'), slide('r4')];

describe('custom shows dialog controller', () => {
	it('creates, renames, copies and removes shows with numeric ids', () => {
		const c = createCustomShowsController([]);
		c.create('Short');
		c.rename('Brief');
		c.copySelected();
		const names = c.getSnapshot().shows.map((s) => [s.id, s.name]);
		expect(names).toEqual([
			['0', 'Brief'],
			['1', 'Copy of Brief'],
		]);
		expect(c.getSnapshot().selectedId).toBe('1');
		c.removeSelected();
		expect(c.getSnapshot().shows).toHaveLength(1);
		expect(c.getSnapshot().selectedId).toBe('0');
		expect(c.getSnapshot().removedIds).toEqual(['1']);
	});

	it('adds the selected presentation slides in deck order and removes show entries', () => {
		const c = createCustomShowsController([]);
		c.create('A');
		c.toggleLeft('r3');
		c.toggleLeft('r1');
		c.addSelected(slides);
		expect(c.getSnapshot().shows[0]?.slideRIds).toEqual(['r1', 'r3']);
		c.addSlide('r2');
		c.toggleRight(0);
		c.removeSelectedSlides();
		expect(c.getSnapshot().shows[0]?.slideRIds).toEqual(['r3', 'r2']);
	});

	it('moves a selected block up and down and reports the limits', () => {
		const c = createCustomShowsController([
			{ id: '0', name: 'A', slideRIds: ['r1', 'r2', 'r3', 'r4'] },
		]);
		c.toggleRight(2);
		c.toggleRight(3);
		expect(c.view(slides).canMoveDown).toBe(false);
		c.moveUp();
		expect(c.getSnapshot().shows[0]?.slideRIds).toEqual(['r1', 'r3', 'r4', 'r2']);
		expect(c.getSnapshot().rightSelected).toEqual([1, 2]);
		c.moveUp();
		c.moveUp();
		expect(c.getSnapshot().shows[0]?.slideRIds).toEqual(['r3', 'r4', 'r1', 'r2']);
		expect(c.view(slides).canMoveUp).toBe(false);
		c.moveDown();
		expect(c.getSnapshot().shows[0]?.slideRIds).toEqual(['r1', 'r3', 'r4', 'r2']);
	});

	it('clears jumps into removed shows on commit, not into a reused id', () => {
		const jump = (id: string) => ({
			id: 'e',
			type: 'shape',
			actionClick: { action: `ppaction://customshow?id=${id}` },
		});
		const deck = [slide('r1', [jump('0')]), slide('r2', [jump('1')])];
		const c = createCustomShowsController([
			{ id: '0', name: 'A', slideRIds: [] },
			{ id: '1', name: 'B', slideRIds: [] },
		]);
		c.select('0');
		c.removeSelected();
		const result = c.result(deck);
		expect(result.shows.map((s) => s.id)).toEqual(['1']);
		const [first, second] = result.slides as [PptxSlide, PptxSlide];
		expect((first.elements[0] as PptxElement).actionClick).toBeUndefined();
		expect((second.elements[0] as PptxElement).actionClick).toBeDefined();
		expect(createCustomShowsController([]).result(deck).slides).toBeUndefined();
	});

	it('describes both lists for the template', () => {
		const withTitle = [slide('r1', [{ type: 'text', text: 'Hello\nworld' }]), slide('r2')];
		const c = createCustomShowsController([{ id: '0', name: 'A', slideRIds: ['r2', 'gone'] }]);
		const view = c.view(withTitle);
		expect(view.presentationSlides.map((r) => [r.number, r.title])).toEqual([
			[1, 'Hello'],
			[2, ''],
		]);
		expect(view.showSlides.map((r) => r.rId)).toEqual(['r2']);
		expect(view.canEdit).toBe(true);
	});
});
