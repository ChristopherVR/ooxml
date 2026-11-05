import { signal } from '@angular/core';
import type { PptxElement, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTransitionOverlayState } from './presentation-transition-overlay-state';

/**
 * A shape whose flip changes turns over in 3D under a camera over the SLIDE's
 * centre (shared `morph-flip-turn`), so the overlay hands its canvas size to
 * the plan: a 960 px wide slide puts the eye (960 / 2) / tan(30 deg) =
 * 831.384 px away, where the 16:9 default would put it at 1108.513 px.
 */
function triangle(slideId: string, flipHorizontal: boolean): PptxSlide {
	return {
		id: slideId,
		elements: [
			{
				id: `${slideId}-tri`,
				type: 'shape',
				name: '!!Tri',
				shapeType: 'rtTriangle',
				x: 100,
				y: 100,
				width: 300,
				height: 200,
				flipHorizontal,
			},
		],
	} as unknown as PptxSlide;
}

describe('createTransitionOverlayState slide size', () => {
	it('sees a turning flip from the centre of its canvas', () => {
		const state = createTransitionOverlayState({
			outgoingSlide: signal(triangle('out', false)),
			incomingSlide: signal(triangle('in', true)),
			canvasSize: signal({ width: 960, height: 540 }),
			transition: signal({ type: 'morph', duration: 800 } as PptxSlideTransition),
			templateElements: signal<readonly PptxElement[]>([]),
			durationMs: signal<number | undefined>(800),
			zoom: signal(1),
		});
		expect(state.morphPlan()?.keyframesCss).toContain('perspective(831.384px)');
	});
});
