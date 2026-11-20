/**
 * Closing a ribbon drop-down menu (More actions, Presentation options) from the keyboard.
 *
 * Each binding opened these menus with its own state and only some of them listened for
 * Escape, so the same menu closed on Escape in one framework and stayed open in another. The
 * listener is shared: a binding calls {@link listenForPopupMenuDismiss} when its menu opens and
 * the returned function when it closes.
 *
 * @module render/popup-menu-dismiss
 */

import { isOutlineViewDismissKey } from './outline-view-dismiss';

/** What {@link listenForPopupMenuDismiss} needs from the binding. */
export interface PopupMenuDismissOptions {
	/** Close the menu (set the binding's open state to false). */
	onDismiss: () => void;
	/** The button that opened the menu: focus goes back to it, as in Office. */
	trigger?: HTMLElement | null | undefined;
}

/**
 * Close an open menu on a bare Escape wherever focus is. Listens on the document in the capture
 * phase, so the key is seen with focus on the trigger or inside the menu, and stops it before
 * the editor's own Escape (clear the selection) runs. Returns the detach function.
 */
export function listenForPopupMenuDismiss(
	doc: Document,
	options: PopupMenuDismissOptions,
): () => void {
	const onKeyDown = (event: KeyboardEvent): void => {
		// The same key rule as Outline view: bare Escape, never one cancelling an IME composition.
		if (!isOutlineViewDismissKey(event)) {
			return;
		}
		event.preventDefault();
		event.stopPropagation();
		options.onDismiss();
		options.trigger?.focus();
	};
	doc.addEventListener('keydown', onKeyDown, true);
	return () => doc.removeEventListener('keydown', onKeyDown, true);
}
