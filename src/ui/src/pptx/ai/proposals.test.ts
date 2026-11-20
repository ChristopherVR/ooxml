import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import type { PptxAiSlidesUpdater } from './bridge';
import { makeMockBridge, makeSlide, textElement } from './mock-bridge';
import { ProposalStore } from './proposals';

const addSlideUpdater: PptxAiSlidesUpdater = (slides: PptxSlide[]) => {
	slides.push(makeSlide(slides.length, [textElement('new-el', 'New slide')]));
	return slides;
};

describe('proposal store', () => {
	it('stages a change with a diff summary without applying it', () => {
		const bridge = makeMockBridge();
		const store = new ProposalStore(bridge);
		const proposal = store.stage('Add slide', addSlideUpdater);

		expect(store.size).toBe(1);
		expect(bridge.edits).toHaveLength(0);
		expect(proposal.summary.join(' ')).toContain('Add 1 slide');
		expect(store.get(proposal.id)?.label).toBe('Add slide');
	});

	it('applies a staged proposal as one history entry', () => {
		const bridge = makeMockBridge();
		const store = new ProposalStore(bridge);
		const before = bridge.getSlides().length;
		const proposal = store.stage('Add slide', addSlideUpdater);

		expect(store.apply(proposal.id)).toBeTruthy();
		expect(bridge.edits).toHaveLength(1);
		expect(bridge.getSlides()).toHaveLength(before + 1);
		expect(store.size).toBe(0);
	});

	it('reverts a staged proposal without touching the deck', () => {
		const bridge = makeMockBridge();
		const store = new ProposalStore(bridge);
		const proposal = store.stage('Add slide', addSlideUpdater);

		expect(store.revert(proposal.id)).toBeTruthy();
		expect(store.size).toBe(0);
		expect(bridge.edits).toHaveLength(0);
	});

	it('applies every staged proposal with acceptAll', () => {
		const bridge = makeMockBridge();
		const store = new ProposalStore(bridge);
		store.stage('Add slide A', addSlideUpdater);
		store.stage('Add slide B', addSlideUpdater);

		expect(store.acceptAll()).toBe(2);
		expect(bridge.edits).toHaveLength(2);
		expect(store.size).toBe(0);
	});

	it('returns false when applying or reverting an unknown id', () => {
		const store = new ProposalStore(makeMockBridge());
		expect(store.apply('nope')).toBeFalsy();
		expect(store.revert('nope')).toBeFalsy();
	});

	it('drops what was staged on a deck once another one loads, and says so', () => {
		let load = 0;
		const bridge = { ...makeMockBridge(), getDeckLoad: () => load };
		const store = new ProposalStore(bridge);
		const staged = store.stage('Add slide', addSlideUpdater);
		const slides = bridge.getSlides();

		load = 1;

		expect(store.list()).toEqual([]);
		expect(store.discarded).toBe(true);
		expect(store.apply(staged.id)).toBeFalsy();
		expect(store.acceptAll()).toBe(0);
		expect(bridge.getSlides()).toBe(slides);
		store.stage('Add slide on the new deck', addSlideUpdater);
		expect(store.size).toBe(1);
		expect(store.discarded).toBe(false);
	});

	it('keeps proposals applicable when an edit replaces the handler', () => {
		let handler = {};
		const bridge = { ...makeMockBridge(), getHandler: () => handler as never };
		const store = new ProposalStore(bridge);
		const staged = store.stage('Add slide', addSlideUpdater);
		const before = bridge.getSlides().length;

		handler = {};

		expect(store.list().map((view) => view.id)).toEqual([staged.id]);
		expect(store.discarded).toBe(false);
		expect(store.apply(staged.id)).toBe(true);
		expect(bridge.getSlides()).toHaveLength(before + 1);
	});

	it('reports nothing discarded when a load finds nothing staged', () => {
		let load = 0;
		const store = new ProposalStore({ ...makeMockBridge(), getDeckLoad: () => load });

		load = 1;

		expect(store.discarded).toBe(false);
	});

	it('tells a store rebuilt on the same bridge what a load dropped', () => {
		let load = 0;
		const bridge = { ...makeMockBridge(), getDeckLoad: () => load };
		new ProposalStore(bridge).stage('Add slide', addSlideUpdater);

		load = 1;
		const rebuilt = new ProposalStore(bridge);

		expect(rebuilt.list()).toEqual([]);
		expect(rebuilt.discarded).toBe(true);
	});

	it('keeps staged proposals for a store rebuilt on the same load', () => {
		const bridge = makeMockBridge();
		const staged = new ProposalStore(bridge).stage('Add slide', addSlideUpdater);

		expect(new ProposalStore(bridge).list().map((view) => view.id)).toEqual([staged.id]);
	});
});
