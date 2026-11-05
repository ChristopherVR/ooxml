import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { useMorphTransitionOverlay } from './use-morph-transition-overlay';

/**
 * A shape whose flip changes turns over in 3D under a camera over the SLIDE's
 * centre (shared `morph-flip-turn`), so the overlay hands the slide size to
 * the plan: a 4:3 slide (960 px wide) puts the eye (960 / 2) / tan(30 deg)
 * = 831.384 px away, where the 16:9 default would put it at 1108.513 px.
 */
function triangle(id: string, flipHorizontal: boolean): PptxElement {
	return {
		id,
		name: '!!Tri',
		type: 'shape',
		shapeType: 'rtTriangle',
		x: 100,
		y: 100,
		width: 300,
		height: 200,
		flipHorizontal,
	} as unknown as PptxElement;
}

const transition = { type: 'morph', duration: 2000 } as PptxSlideTransition;
const from = { id: 'a', elements: [triangle('a-1', false)] } as unknown as PptxSlide;
const to = { id: 'b', elements: [triangle('b-1', true)] } as unknown as PptxSlide;

describe('useMorphTransitionOverlay slide size', () => {
	it('sees a turning flip from the slide it is given', () => {
		const morph = useMorphTransitionOverlay({
			transition: () => transition,
			outgoingSlide: () => from,
			incomingSlide: () => to,
			durationMs: () => 2000,
			slideSize: () => ({ width: 960, height: 720 }),
		});
		expect(morph.morphPlan.value?.keyframesCss).toContain('perspective(831.384px)');
	});
});
