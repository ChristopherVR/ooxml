import { createBroadcastFollower } from 'ooxml-ui/pptx';

import type { Store, ViewerState } from '../state';
import type { CollaborationController } from './collaboration-controller-types';

/**
 * Keep the session's presence and the local slide in step, off the store, so no
 * other module needs to know about collaboration to stay presence-aware:
 * publish the local active slide and selection, follow a peer the user chose
 * to follow, and keep a broadcast's audience on the presenter's slide (the
 * shared rule: only a local `viewer` follows the `owner`, each time it moves).
 * Returns the unsubscribe.
 */
export function syncPresenceNavigation(
	store: Store<ViewerState>,
	collaboration: CollaborationController,
	goToSlide: (index: number) => void,
): () => void {
	const broadcast = createBroadcastFollower();
	return store.subscribe((state, previous) => {
		if (!collaboration.isActive()) {
			return;
		}
		if (state.currentSlide !== previous.currentSlide) {
			collaboration.setActiveSlide(state.currentSlide);
		}
		if (state.selectedElementId !== previous.selectedElementId) {
			collaboration.setSelection(state.selectedElementId ?? undefined, state.currentSlide);
		}
		if (state.followedClientId !== null && state.remotePresences !== previous.remotePresences) {
			const followed = state.remotePresences.find((p) => p.clientId === state.followedClientId);
			if (followed && followed.activeSlideIndex !== state.currentSlide) {
				goToSlide(followed.activeSlideIndex);
			}
		}
		// The deck the audience joined with arrives after the presenter's presence can.
		if (state.remotePresences === previous.remotePresences && state.slides === previous.slides) {
			return;
		}
		const presenterSlide = broadcast.step({
			localRole: collaboration.getConfig()?.role,
			peers: state.remotePresences,
			slideCount: state.slides.length,
			followingPeer: state.followedClientId !== null,
			currentSlide: state.currentSlide,
		});
		if (presenterSlide !== null) {
			goToSlide(presenterSlide);
		}
	});
}
