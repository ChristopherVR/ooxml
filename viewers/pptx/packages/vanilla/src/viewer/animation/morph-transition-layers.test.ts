import type { MorphTransitionPlan } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphCrossfadeGroups, buildMorphOverlayLayers } from './morph-transition-layers';

function stage(ids: string[]): HTMLElement {
	const root = document.createElement('div');
	root.className = 'pptxv-stage';
	ids.forEach((id, index) => {
		const box = document.createElement('div');
		box.dataset.elementId = id;
		// What the renderer gives each shape: its place in ITS OWN slide.
		box.style.zIndex = String(index + 1);
		root.appendChild(box);
	});
	return root;
}

describe('buildMorphOverlayLayers', () => {
	// Regression (issue #141, slides 9 to 10): the left column's ghost ranks
	// above the full-slide backdrop's ghost in the plan, but sat under it on the
	// outgoing slide, and the cloned stage kept that order, hiding the column.
	it('stacks a run in the order the plan ranked it, not the slide it was cloned from', () => {
		const plan = {
			overlayLayers: [{ side: 'outgoing', elements: [{ id: 'backdrop' }, { id: 'column' }] }],
		} as unknown as MorphTransitionPlan;
		const [layer] = buildMorphOverlayLayers(
			document,
			stage(['column', 'backdrop']),
			stage([]),
			plan,
		);
		const z = (id: string): number =>
			Number(layer!.querySelector<HTMLElement>(`[data-element-id="${id}"]`)!.style.zIndex);
		expect(z('column')).toBeGreaterThan(z('backdrop'));
	});

	it('paints an offstage ghost with its destination text layout', () => {
		const outgoing = stage(['old']);
		const old = outgoing.querySelector<HTMLElement>('[data-element-id="old"]')!;
		old.style.left = '-500px';
		old.style.top = '40px';
		old.style.height = '100px';
		old.textContent = 'old layout';
		const incoming = stage(['new']);
		const target = incoming.querySelector<HTMLElement>('[data-element-id="new"]')!;
		target.style.left = '300px';
		target.style.top = '60px';
		target.style.height = '110px';
		target.textContent = 'new layout';
		const plan = {
			overlayLayers: [{ side: 'outgoing', elements: [{ id: 'old' }] }],
			outgoingTextReplacements: new Map([['old', 'new']]),
		} as unknown as MorphTransitionPlan;
		const [layer] = buildMorphOverlayLayers(document, outgoing, incoming, plan);
		const ghost = layer!.querySelector<HTMLElement>('[data-element-id="old"]')!;
		expect(ghost.textContent).toBe('new layout');
		expect([ghost.style.left, ghost.style.top, ghost.style.height]).toEqual([
			'-500px',
			'40px',
			'110px',
		]);
		expect(old.textContent).toBe('old layout');
		expect(target.dataset.elementId).toBe('new');
	});
});

describe('buildMorphCrossfadeGroups', () => {
	// A rewrapped text pair is painted as one copy of each half per piece
	// (`textPieceSources`). The slide never rendered those copies, so the
	// stage clone has to grow them from the element each one copies.
	it('paints a text piece as a clone of the element it copies', () => {
		const outgoing = stage(['old']);
		outgoing.querySelector<HTMLElement>('[data-element-id="old"]')!.textContent = 'old words';
		const incoming = stage(['new']);
		incoming.querySelector<HTMLElement>('[data-element-id="new"]')!.textContent = 'new words';
		const plan = {
			overlayLayers: [],
			crossfadeGroups: [
				{
					outgoing: { id: 'old-morph-piece-0', height: 40 },
					incoming: { id: 'new-morph-piece-0', height: 60 },
				},
				{
					outgoing: { id: 'old-morph-piece-1', height: 40 },
					incoming: { id: 'new-morph-piece-1', height: 60 },
				},
			],
			textPieceSources: new Map([
				['old-morph-piece-0', 'old'],
				['old-morph-piece-1', 'old'],
				['new-morph-piece-0', 'new'],
				['new-morph-piece-1', 'new'],
			]),
		} as unknown as MorphTransitionPlan;
		const groups = buildMorphCrossfadeGroups(document, outgoing, incoming, plan);
		expect(groups).toHaveLength(2);
		const ids = (group: HTMLElement) =>
			[...group.querySelectorAll<HTMLElement>('[data-element-id]')].map((node) => [
				node.dataset.elementId,
				node.textContent,
			]);
		expect(ids(groups[1]!)).toEqual([
			['old-morph-piece-1', 'old words'],
			['new-morph-piece-1', 'new words'],
		]);
		// Grown to the copy's box, whose centre its keyframes pivot on.
		const copy = groups[1]!.querySelector<HTMLElement>('[data-element-id="new-morph-piece-1"]')!;
		expect(copy.style.height).toBe('60px');
		// The source stages are cloned, never modified.
		expect(outgoing.querySelectorAll('[data-element-id]')).toHaveLength(1);
	});
});
