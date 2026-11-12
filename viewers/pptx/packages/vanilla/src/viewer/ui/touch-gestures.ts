import { createTouchGestureRecognizer, MAX_ZOOM_SCALE, MIN_ZOOM_SCALE } from 'ooxml-ui/pptx';

export interface TouchGestureHandlers {
	/** Read the current zoom factor relative to fit (1 = fit), as the zoom buttons do. */
	getScale(): number;
	/**
	 * Persist a fit-relative zoom after a two-finger pinch, already clamped to the
	 * shared bounds the buttons, Ctrl+wheel and `setZoom` use.
	 */
	onPinchZoom(scale: number): void;
	/** Whether a one-finger swipe may change slides at this moment. */
	isSwipeEnabled(): boolean;
	onNext(): void;
	onPrevious(): void;
}

/**
 * Attach React-parity touch gestures to the stable viewport node. Every pinch
 * reaches `onPinchZoom` (the viewer ignores it during a slide show); slide
 * swipes deliberately stay disabled while editing so they cannot steal element
 * move/resize gestures.
 */
export function attachTouchGestures(
	target: HTMLElement,
	handlers: TouchGestureHandlers,
): () => void {
	const recognizer = createTouchGestureRecognizer({
		getScale: handlers.getScale,
		minScale: MIN_ZOOM_SCALE,
		maxScale: MAX_ZOOM_SCALE,
		callbacks: {
			onPinchZoom: handlers.onPinchZoom,
			onSwipe: (direction) => {
				if (!handlers.isSwipeEnabled()) {
					return;
				}
				if (direction === -1) {
					handlers.onNext();
				} else {
					handlers.onPrevious();
				}
			},
		},
	});

	const onStart = (event: TouchEvent): void => recognizer.onTouchStart(event);
	const onMove = (event: TouchEvent): void => recognizer.onTouchMove(event);
	const onEnd = (event: TouchEvent): void => recognizer.onTouchEnd(event);
	const onCancel = (): void => recognizer.onTouchCancel();
	target.addEventListener('touchstart', onStart, { passive: false, capture: true });
	target.addEventListener('touchmove', onMove, { passive: false, capture: true });
	target.addEventListener('touchend', onEnd, { passive: true, capture: true });
	target.addEventListener('touchcancel', onCancel, { passive: true, capture: true });

	return () => {
		target.removeEventListener('touchstart', onStart, true);
		target.removeEventListener('touchmove', onMove, true);
		target.removeEventListener('touchend', onEnd, true);
		target.removeEventListener('touchcancel', onCancel, true);
		recognizer.cancel();
	};
}
