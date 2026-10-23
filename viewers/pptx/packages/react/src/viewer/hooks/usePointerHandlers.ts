/**
 * usePointerHandlers: the canvas's pointer gestures.
 *
 *   - move and resize run on the shared gesture controller
 *     (canvas-gestures.ts), created once here and handed to the press handlers
 *     through `gestureRef`;
 *   - the marquee and the shape-adjustment drags keep the global listeners
 *     below (pointer-move-handlers.ts, pointer-up-handlers.ts).
 */
import { useEffect, useMemo, useRef } from 'react';

import { createCanvasGestures } from './canvas-gestures';
import type { PointerFrameTracker, UsePointerHandlersInput } from './pointer-handler-types';
import { processPointerMove } from './pointer-move-handlers';
import { processPointerUp } from './pointer-up-handlers';

export type { UsePointerHandlersInput };

export function usePointerHandlers(input: UsePointerHandlersInput): void {
	const {
		editorScale,
		canvasStageRef,
		canvasSize,
		activeSlide,
		activeSlideIndex,
		shapeAdjustmentDragStateRef,
		marqueeStateRef,
		editTemplateMode,
		templateElements,
		setMarqueeSelectionState,
		setPointerCommitNonce,
		applySelection,
		clearSelection,
		updateElementById,
		markDirty,
		livePatcher,
		gestureRef,
	} = input;

	// The gesture controller lives for the whole mount and reads the latest
	// input on every frame, so a re-render mid-drag never tears it down.
	const inputRef = useRef(input);
	inputRef.current = input;
	const gestures = useMemo(() => createCanvasGestures(() => inputRef.current), []);
	useEffect(() => {
		if (gestureRef) {
			gestureRef.current = gestures;
		}
		return () => {
			gestures.dispose();
			if (gestureRef?.current === gestures) {
				gestureRef.current = null;
			}
		};
	}, [gestures, gestureRef]);

	// The effect body closes over `input` as a whole (passed straight through to
	// processPointerMove/processPointerUp) rather than the destructured fields
	// above, so they never show up as "read" to the dependency analyzer. They are
	// listed below anyway, deliberately, so the global listeners only
	// re-subscribe when one of them actually changes. `input` itself is a fresh
	// object every render; listing it would rebuild the listeners on every
	// render and cancel a marquee frame scheduled mid-render.
	/* oxlint-disable react-hooks/exhaustive-deps -- see comment above */
	useEffect(() => {
		const tracker: PointerFrameTracker = { rafId: 0, pendingMoveEvent: null };

		const handlePointerMove = (e: PointerEvent) => {
			tracker.pendingMoveEvent = e;
			if (tracker.rafId === 0) {
				tracker.rafId = requestAnimationFrame(() => {
					tracker.rafId = 0;
					if (tracker.pendingMoveEvent) {
						processPointerMove(tracker.pendingMoveEvent, input);
						tracker.pendingMoveEvent = null;
					}
				});
			}
		};

		const handlePointerUp = () => {
			if (tracker.rafId !== 0) {
				cancelAnimationFrame(tracker.rafId);
				tracker.rafId = 0;
				tracker.pendingMoveEvent = null;
			}
			processPointerUp(input);
		};

		document.addEventListener('pointermove', handlePointerMove);
		document.addEventListener('pointerup', handlePointerUp);
		document.addEventListener('pointercancel', handlePointerUp);
		return () => {
			document.removeEventListener('pointermove', handlePointerMove);
			document.removeEventListener('pointerup', handlePointerUp);
			document.removeEventListener('pointercancel', handlePointerUp);
			if (tracker.rafId !== 0) {
				cancelAnimationFrame(tracker.rafId);
			}
		};
	}, [
		editorScale,
		canvasStageRef,
		shapeAdjustmentDragStateRef,
		marqueeStateRef,
		editTemplateMode,
		templateElements,
		setMarqueeSelectionState,
		setPointerCommitNonce,
		activeSlide,
		activeSlideIndex,
		canvasSize.width,
		canvasSize.height,
		applySelection,
		clearSelection,
		updateElementById,
		markDirty,
		livePatcher,
		gestureRef,
	]);
	/* oxlint-enable react-hooks/exhaustive-deps */
}
