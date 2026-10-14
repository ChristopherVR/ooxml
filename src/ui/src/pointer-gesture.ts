/** Capture can transfer from a touch's implicit child target to the gesture owner. */
export function ownsCaptureLoss(owner: Element, event: PointerEvent): boolean {
	return event.target === owner;
}

/** Scroll an editor viewport from one finger on empty canvas. */
export function trackPointerPan(
	owner: HTMLElement,
	start: PointerEvent,
	onMove: () => void,
	onEnd: () => void,
): () => void {
	const left = owner.scrollLeft,
		top = owner.scrollTop;
	return trackPointerGesture(owner, start.pointerId, {
		move: (event) => {
			if (Math.hypot(start.clientX - event.clientX, start.clientY - event.clientY) < 4) return;
			owner.scrollLeft = left + start.clientX - event.clientX;
			owner.scrollTop = top + start.clientY - event.clientY;
			onMove();
		},
		commit: onEnd,
		cancel: onEnd,
	});
}

/** One pointer, one commit; cancellation restores the preview and never commits. */
export function trackPointerGesture(
	owner: HTMLElement,
	pointerId: number,
	callbacks: {
		move(event: PointerEvent): void;
		commit(event: PointerEvent): void;
		cancel(): void;
	},
): () => void {
	const win = owner.ownerDocument.defaultView;
	const Abort = win?.AbortController ?? AbortController;
	const events = new Abort();
	const options = { signal: events.signal };
	let active = true;
	const finish = (event?: PointerEvent) => {
		if (!active) return;
		active = false;
		events.abort();
		if (owner.hasPointerCapture?.(pointerId)) owner.releasePointerCapture(pointerId);
		if (event) callbacks.commit(event);
		else callbacks.cancel();
	};
	owner.addEventListener(
		'pointermove',
		(event) => {
			if (event.pointerId === pointerId) callbacks.move(event);
		},
		options,
	);
	owner.addEventListener(
		'pointerup',
		(event) => {
			if (event.pointerId === pointerId) finish(event);
		},
		options,
	);
	owner.addEventListener(
		'pointercancel',
		(event) => {
			if (event.pointerId === pointerId) finish();
		},
		options,
	);
	owner.addEventListener(
		'lostpointercapture',
		(event) => {
			if (event.pointerId === pointerId && ownsCaptureLoss(owner, event)) finish();
		},
		options,
	);
	owner.ownerDocument.addEventListener(
		'keydown',
		(event) => {
			if (event.key === 'Escape') {
				event.preventDefault();
				finish();
			}
		},
		options,
	);
	win?.addEventListener('blur', () => finish(), options);
	owner.setPointerCapture?.(pointerId);
	return () => finish();
}
