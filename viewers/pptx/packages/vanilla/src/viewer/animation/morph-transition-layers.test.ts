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

	// A shape whose text is carried apart from its body (`morphCopyParts`) is
	// painted as a copy without its text and a copy without its fill and outline.
	it('strips a body or text copy down to its part', () => {
		const shape = (id: string) => {
			const root = stage([id]);
			const node = root.querySelector<HTMLElement>(`[data-element-id="${id}"]`)!;
			node.style.backgroundColor = 'rgb(32, 112, 192)';
			node.innerHTML =
				'<svg class="outline"></svg><div class="pptxv-text"><p>Lorem ipsum</p></div>';
			return root;
		};
		const plan = {
			overlayLayers: [],
			crossfadeGroups: [
				{
					outgoing: { id: 'a-morph-body', height: 160 },
					incoming: { id: 'b-morph-body', height: 240 },
				},
				{
					outgoing: { id: 'a-morph-text', height: 160 },
					incoming: { id: 'b-morph-text', height: 240 },
				},
			],
			textPieceSources: new Map([
				['a-morph-body', 'a'],
				['a-morph-text', 'a'],
				['b-morph-body', 'b'],
				['b-morph-text', 'b'],
			]),
			morphCopyParts: new Map([
				['a-morph-body', 'body'],
				['a-morph-text', 'text'],
				['b-morph-body', 'body'],
				['b-morph-text', 'text'],
			]),
		} as unknown as MorphTransitionPlan;
		const [body, text] = buildMorphCrossfadeGroups(document, shape('a'), shape('b'), plan);
		const copy = (group: HTMLElement | undefined, id: string) =>
			group!.querySelector<HTMLElement>(`[data-element-id="${id}"]`)!;
		const bodyCopy = copy(body, 'b-morph-body');
		expect(bodyCopy.querySelector('.pptxv-text')).toBeNull();
		expect(bodyCopy.querySelector('svg')).not.toBeNull();
		expect(bodyCopy.style.backgroundColor).toBe('rgb(32, 112, 192)');
		const textCopy = copy(text, 'a-morph-text');
		expect(textCopy.textContent).toBe('Lorem ipsum');
		expect(textCopy.querySelector('svg')).toBeNull();
		expect(textCopy.style.backgroundColor).not.toBe('rgb(32, 112, 192)');
	});
});
