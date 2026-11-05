/**
 * The slide-transition overlay as a running show mounts it: one overlay per
 * transition, behind each slide's own layout.
 *
 * Split out of `ViewerCanvasArea` so the two decisions it makes are tested on
 * the code the show runs, not on a copy of it:
 *
 *  - The KEY. The overlay re-arms its completion timer only when its duration
 *    or `onComplete` change, and the show's `onComplete` is stable. Two
 *    transitions of the same length in a row would otherwise share one
 *    overlay, and the first one's timer would take the second down partway
 *    through (issue #141). Keying by the slide pair mounts a fresh overlay,
 *    and so a fresh timer, for every transition.
 *  - The TEMPLATE of each layer, from `transitionTemplateElements`: the
 *    leaving slide over its own layout, not the arriving slide's (issue #141,
 *    slides 1 to 2).
 */
import { transitionTemplateElements } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import React from 'react';

import type { UsePresentationModeResult } from '../hooks/usePresentationMode';
import type { CanvasSize } from '../types';
import { PresentationTransitionOverlay } from './PresentationTransitionOverlay';

export interface PresentationTransitionMountProps {
	presentation: Pick<
		UsePresentationModeResult,
		'transitionOverlay' | 'morphPlan' | 'handleTransitionOverlayComplete'
	>;
	slides: readonly PptxSlide[];
	templateElementsBySlideId: Readonly<Record<string, readonly PptxElement[]>>;
	canvasSize: CanvasSize;
	/** The stage's scale, so the overlay matches the live slide box. */
	scale: number;
}

export function PresentationTransitionMount({
	presentation,
	slides,
	templateElementsBySlideId,
	canvasSize,
	scale,
}: PresentationTransitionMountProps): React.ReactElement | null {
	const overlay = presentation.transitionOverlay;
	const outgoingSlide = overlay ? slides[overlay.outgoingSlideIndex] : undefined;
	if (!overlay || !outgoingSlide) {
		return null;
	}
	const incomingSlide = slides[overlay.incomingSlideIndex];
	const templates = transitionTemplateElements(
		outgoingSlide,
		incomingSlide,
		templateElementsBySlideId,
	);
	return (
		<PresentationTransitionOverlay
			key={`${overlay.outgoingSlideIndex}-${overlay.incomingSlideIndex}`}
			outgoingSlide={outgoingSlide}
			templateElements={templates.outgoing}
			incomingTemplateElements={templates.incoming}
			canvasSize={canvasSize}
			transition={overlay.transition}
			durationMs={overlay.durationMs}
			scale={scale}
			morphPlan={presentation.morphPlan}
			incomingSlide={incomingSlide}
			onComplete={presentation.handleTransitionOverlayComplete}
		/>
	);
}
