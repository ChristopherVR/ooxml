import { externalContentBarVisible } from 'ooxml-ui/pptx';

import type { LoadingController } from './loading-controller';
import type { Store, ViewerState } from './state';

/**
 * PowerPoint's security bar for the loaded deck: it links pictures or media to
 * web addresses, the load left them out, the host has not blocked them
 * outright and the user has not closed the bar.
 */
export function externalContentBarShown(state: ViewerState, host: boolean | undefined): boolean {
	return externalContentBarVisible({
		hasLinkedWebContent: state.hasLinkedWebContent,
		allowed: state.externalContentAllowed,
		host,
		closed: state.externalContentBarClosed,
	});
}

export interface EnableExternalContentDeps {
	store: Store<ViewerState>;
	loading: LoadingController;
	/** The current deck with its unsaved changes, serialised as a recovery snapshot is. */
	serializeForRecovery(): Promise<Uint8Array | null>;
}

/**
 * Enable Content: reopen this presentation with its linked pictures and media.
 * Without unsaved changes it reloads the bytes it was opened from; with them,
 * the deck as it stands (still marked unsaved afterwards). The slide on screen
 * stays, and so does Protected View when it is on. Undo history starts over,
 * as with any load.
 */
export async function enableExternalContent(deps: EnableExternalContentDeps): Promise<void> {
	const { store, loading } = deps;
	const before = store.get();
	// The bar goes at once; the reload decides what shows next.
	store.set({ externalContentBarClosed: true });
	const unsaved = before.dirty;
	const bytes = unsaved ? await deps.serializeForRecovery() : loading.getLoadedBytes();
	if (!bytes) {
		return;
	}
	await loading.load(bytes, 'user', {
		skipProtectedView: !before.protectedView,
		enableExternalContent: true,
	});
	const after = store.get();
	if (after.error !== null || after.slides.length === 0) {
		return;
	}
	store.set({
		currentSlide: Math.min(before.currentSlide, after.slides.length - 1),
		...(unsaved ? { dirty: true } : {}),
	});
}
