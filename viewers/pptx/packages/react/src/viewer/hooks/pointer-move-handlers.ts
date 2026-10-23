import { getDraggedShapeAdjustments } from 'ooxml-ui/pptx';
/**
 * Pointer-move processing for usePointerHandlers: the marquee and the
 * shape-adjustment drags. Move and resize run on the shared gesture controller
 * (see canvas-gestures.ts).
 */
import type { PptxElement } from 'pptx-viewer-core';

import type { UsePointerHandlersInput } from './pointer-handler-types';

// ---------------------------------------------------------------------------
// Main pointer-move processor
// ---------------------------------------------------------------------------

export function processPointerMove(e: PointerEvent, input: UsePointerHandlersInput): void {
	const {
		editorScale,
		canvasStageRef,
		canvasSize,
		marqueeStateRef,
		shapeAdjustmentDragStateRef,
		setMarqueeSelectionState,
		updateElementById,
	} = input;

	const marquee = marqueeStateRef.current;
	if (marquee) {
		processMarqueeMove(
			e,
			marquee,
			canvasStageRef,
			canvasSize,
			editorScale,
			setMarqueeSelectionState,
		);
		return;
	}

	const adj = shapeAdjustmentDragStateRef.current;
	if (adj) {
		processAdjustmentMove(e, adj, editorScale, updateElementById);
	}
}

// ── Marquee ──────────────────────────────────────────────────────────────────

function processMarqueeMove(
	e: PointerEvent,
	marquee: NonNullable<UsePointerHandlersInput['marqueeStateRef']['current']>,
	canvasStageRef: UsePointerHandlersInput['canvasStageRef'],
	canvasSize: UsePointerHandlersInput['canvasSize'],
	editorScale: number,
	setMarqueeSelectionState: UsePointerHandlersInput['setMarqueeSelectionState'],
): void {
	const stage = canvasStageRef.current;
	if (!stage) {
		return;
	}
	const rect = stage.getBoundingClientRect();
	marquee.currentX = Math.max(0, Math.min(canvasSize.width, (e.clientX - rect.left) / editorScale));
	marquee.currentY = Math.max(0, Math.min(canvasSize.height, (e.clientY - rect.top) / editorScale));
	setMarqueeSelectionState({ ...marquee });
}

// ── Shape adjustment ─────────────────────────────────────────────────────────

function processAdjustmentMove(
	e: PointerEvent,
	adj: NonNullable<UsePointerHandlersInput['shapeAdjustmentDragStateRef']['current']>,
	editorScale: number,
	updateElementById: UsePointerHandlersInput['updateElementById'],
): void {
	const dx = (e.clientX - adj.startClientX) / editorScale;
	// Both axes: a handle only travels horizontally on a round-rect. An arrow's
	// shaft thickness, a callout's leader line and a pie wedge's sweep all need
	// the vertical component, and feeding 0 pinned them to their start value.
	const dy = (e.clientY - adj.startClientY) / editorScale;
	// Shared owns the adjustment maths for all five bindings. React used to keep
	// a private `computeAdjustmentValue` that clamped the result to 0..1, but an
	// `a:avLst` adjustment is a 0..50000 guide value: any drag therefore collapsed
	// a 16667 corner radius to 1, i.e. to nothing. Its unit tests passed because
	// they asserted the same wrong scale.
	const adjustments = getDraggedShapeAdjustments(adj, dx, dy);
	if (!adj.moved && Math.hypot(dx, dy) > 2) {
		adj.moved = true;
	}
	if (adj.moved) {
		// The WHOLE map, not one key: `shapeAdjustments` is replaced wholesale by
		// `updateElementById`, so writing only the dragged guide would delete the
		// other two on a `quadArrow`.
		updateElementById(adj.elementId, {
			shapeAdjustments: adjustments,
		} as Partial<PptxElement>);
	}
}
