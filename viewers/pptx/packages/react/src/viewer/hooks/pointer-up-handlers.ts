import { computeMarqueeHitIds, mergeAdditiveSelection } from 'ooxml-ui/pptx';
import type { MarqueeElementRect, MarqueeRect as SharedMarqueeRect } from 'ooxml-ui/pptx';
/**
 * Pointer-up (commit) logic for usePointerHandlers: commits a marquee
 * selection and ends a shape-adjustment drag. Move and resize end on the
 * shared gesture controller's own release (see canvas-gestures.ts).
 */
import type { PptxElement } from 'pptx-viewer-core';

import type { UsePointerHandlersInput } from './pointer-handler-types';

// ---------------------------------------------------------------------------
// Re-exported pure helpers (now backed by `pptx-viewer-shared`)
// ---------------------------------------------------------------------------

/** A marquee drag described by its start and current corner (any order). */
export type MarqueeRect = SharedMarqueeRect;

/** An element reduced to its id + bounding box for marquee hit-testing. */
export type ElementRect = MarqueeElementRect;

export { computeMarqueeHitIds, mergeAdditiveSelection };

// ---------------------------------------------------------------------------
// Main pointer-up processor
// ---------------------------------------------------------------------------

export function processPointerUp(input: UsePointerHandlersInput): void {
	const {
		editTemplateMode,
		templateElements,
		activeSlide,
		marqueeStateRef,
		dragStateRef,
		resizeStateRef,
		shapeAdjustmentDragStateRef,
		justInteractedRef,
		gestureRef,
		setMarqueeSelectionState,
		setPointerCommitNonce,
		applySelection,
		clearSelection,
		markDirty,
		livePatcher,
	} = input;

	const marquee = marqueeStateRef.current;
	const adj = shapeAdjustmentDragStateRef.current;

	if (marquee) {
		commitMarquee(
			marquee,
			editTemplateMode,
			templateElements,
			activeSlide,
			applySelection,
			clearSelection,
		);
	}

	// A running move/resize ends on the gesture controller's own (window)
	// pointerup, which fires after this document listener and clears these
	// records itself. One that never started (no controller) is just disarmed.
	if (!gestureRef?.current?.isActive()) {
		dragStateRef.current = null;
		resizeStateRef.current = null;
	}

	// The adjustment drag leaves the pointer over the node it went down on, so
	// the trailing `click` must not read as a second click on a selected element.
	if (adj?.moved) {
		justInteractedRef.current = true;
		livePatcher?.flush();
	}

	marqueeStateRef.current = null;
	shapeAdjustmentDragStateRef.current = null;
	setMarqueeSelectionState(null);

	if (adj?.moved) {
		markDirty();
	}
	// Pointer-down can commit text before arming a marquee or an adjustment.
	// Wake history after clearing that interaction even when the pointer stayed
	// still; its document comparison rejects clicks that changed nothing.
	if (adj || marquee) {
		setPointerCommitNonce((n) => n + 1);
	}
}

// ── Marquee commit ───────────────────────────────────────────────────────────

function commitMarquee(
	marquee: NonNullable<UsePointerHandlersInput['marqueeStateRef']['current']>,
	editTemplateMode: boolean,
	templateElements: PptxElement[],
	activeSlide: UsePointerHandlersInput['activeSlide'],
	applySelection: UsePointerHandlersInput['applySelection'],
	clearSelection: UsePointerHandlersInput['clearSelection'],
): void {
	// In edit-template mode the marquee hit-tests the template store (the only
	// interactive layer); otherwise it runs over the active slide's elements.
	const sourceElements = editTemplateMode ? templateElements : (activeSlide?.elements ?? []);
	const hitIds = computeMarqueeHitIds(marquee, sourceElements);
	if (marquee.additive) {
		const merged = mergeAdditiveSelection(marquee.baseSelectionIds, hitIds);
		if (merged.length > 0) {
			applySelection(merged[0], merged);
		} else {
			clearSelection();
		}
	} else if (hitIds.length > 0) {
		applySelection(hitIds[0], hitIds);
	} else {
		clearSelection();
	}
}
