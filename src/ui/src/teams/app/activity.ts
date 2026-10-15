/** Events that count as the person being at the keyboard. */
const ACTIVITY_EVENTS = ['pointerdown', 'pointermove', 'keydown', 'wheel', 'touchstart'] as const;

/**
 * Reports user input on a document, at most once per `throttleMs`, so the client can reset idleness
 * without handling every pointer move. Returns the function that stops listening.
 */
export function watchActivity(
	target: Pick<Document, 'addEventListener' | 'removeEventListener'>,
	onActivity: () => void,
	throttleMs = 1000,
	now: () => number = Date.now,
): () => void {
	let last = Number.NEGATIVE_INFINITY;
	const handler = (): void => {
		const t = now();
		if (t - last < throttleMs) return;
		last = t;
		onActivity();
	};
	for (const name of ACTIVITY_EVENTS) target.addEventListener(name, handler, { passive: true });
	return () => {
		for (const name of ACTIVITY_EVENTS) target.removeEventListener(name, handler);
	};
}
