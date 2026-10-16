/**
 * Ownership of the dialogs the viewer mounts on `document.body`.
 *
 * Those dialogs live outside the viewer's root, so removing the root does not
 * remove them: a viewer destroyed while one is open would leave it on screen,
 * and its closures would keep the whole viewer reachable. Each dialog
 * registers its close with the viewer's {@link OverlayOwner} (the viewer's
 * dialog `RegionSet`), which runs every registered close on destroy.
 *
 * @module viewer/ui/overlay-owner
 */

/** Runs the registered closes when it goes away (the viewer, on destroy). */
export interface OverlayOwner {
	/** Register `close`; returns the function that unregisters it. */
	own(close: () => void): () => void;
}

/**
 * Register `close` with `owner` and return the close a dialog should use from
 * then on: it unregisters before closing, so a dialog closed by the user is
 * not held until the viewer goes. Without an owner `close` is returned as is.
 */
export function ownOverlay(owner: OverlayOwner | undefined, close: () => void): () => void {
	if (!owner) {
		return close;
	}
	let release: (() => void) | undefined;
	const owned = (): void => {
		release?.();
		release = undefined;
		close();
	};
	release = owner.own(owned);
	return owned;
}
