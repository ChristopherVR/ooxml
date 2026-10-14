/**
 * Leaving Outline view from the keyboard.
 *
 * Outline view covers the whole editor, and until this module the only way
 * back to Normal was its close button: Escape did nothing in any binding. Each
 * binding only listened for keys on the outline rows, and opening the view
 * leaves focus on the ribbon button that opened it (behind the overlay), so a
 * row listener never saw the key anyway.
 *
 * The decision ("which key closes it") and the listener are shared so the five
 * bindings cannot drift: each one calls {@link listenForOutlineViewDismiss}
 * when the view opens and the returned function when it closes.
 *
 * @module render/outline-view-dismiss
 */

/** Keyboard event shape read by {@link isOutlineViewDismissKey}. */
export interface OutlineDismissKeyInput {
	key: string;
	shiftKey?: boolean;
	ctrlKey?: boolean;
	metaKey?: boolean;
	altKey?: boolean;
	/** True while an IME composition is open; its Escape cancels the composition. */
	isComposing?: boolean;
}

/**
 * Whether a key press closes Outline view and returns to Normal view.
 *
 * Bare Escape only. A chord is left to the browser and the host, and an Escape
 * that is cancelling an IME composition in a row must not throw the user out
 * of the view in the middle of typing.
 */
export function isOutlineViewDismissKey(input: OutlineDismissKeyInput): boolean {
	return (
		input.key === 'Escape' &&
		!input.isComposing &&
		!input.ctrlKey &&
		!input.metaKey &&
		!input.altKey &&
		!input.shiftKey
	);
}

/** The part of `Document` the listener needs; a test can pass a fake. */
export interface OutlineDismissTarget {
	addEventListener(
		type: 'keydown',
		listener: (event: KeyboardEvent) => void,
		capture: boolean,
	): void;
	removeEventListener(
		type: 'keydown',
		listener: (event: KeyboardEvent) => void,
		capture: boolean,
	): void;
}

/**
 * Close Outline view on Escape wherever focus is, while it is open.
 *
 * Listens on the document in the CAPTURE phase, like the reading view and the
 * slide sorter, so the key is seen even with focus still on the ribbon button
 * that opened the view, and is stopped before the editor's own Escape (clear
 * the selection) runs underneath the overlay. Returns the detach function.
 */
export function listenForOutlineViewDismiss(
	target: OutlineDismissTarget,
	onDismiss: () => void,
): () => void {
	const onKeyDown = (event: KeyboardEvent): void => {
		if (!isOutlineViewDismissKey(event)) {
			return;
		}
		event.preventDefault();
		event.stopPropagation();
		onDismiss();
	};
	target.addEventListener('keydown', onKeyDown, true);
	return () => target.removeEventListener('keydown', onKeyDown, true);
}
