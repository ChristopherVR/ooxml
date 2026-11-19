import { externalContentBarVisible, prepareExternalContentReload } from 'ooxml-ui/pptx';

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
	/** The current deck with its edits, serialised as a recovery snapshot is. */
	serializeForRecovery(): Promise<Uint8Array | null>;
}

/**
 * Enable Content: reopen this presentation as it stands (the shared
 * `prepareExternalContentReload`) with its linked pictures and media. Unsaved
 * changes stay unsaved, the slide on screen stays, and so do Protected View
 * when it is on and the Protect Presentation badge. Undo history starts over.
 */
export async function enableExternalContent(deps: EnableExternalContentDeps): Promise<void> {
	const { store, loading } = deps;
	const before = store.get();
	const reload = await prepareExternalContentReload({
		serializeCurrentDeck: deps.serializeForRecovery,
		isDirty: before.dirty,
		slideIndex: before.currentSlide,
	});
	if (!reload) {
		return;
	}
	// The bar goes at once; the reload decides what shows next.
	store.set({ externalContentBarClosed: true });
	await loading.load(reload.bytes, 'user', {
		skipProtectedView: !before.protectedView,
		enableExternalContent: true,
	});
	const after = store.get();
	if (after.error !== null || after.slides.length === 0) {
		return;
	}
	store.set({
		currentSlide: Math.min(reload.slideIndex, after.slides.length - 1),
		// The reloaded bytes are the plaintext recovery serialisation.
		isPasswordProtected: before.isPasswordProtected,
		...(reload.keepDirty ? { dirty: true } : {}),
	});
}
