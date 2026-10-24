import type { MorphTransitionPlan } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphOverlayLayers } from './morph-transition-layers';

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
});
