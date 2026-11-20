/**
 * collaboration-presence-effects.svelte.ts: publishes local active-slide and
 * selection changes into the collaboration session, drives follow-mode
 * navigation when the local user is following a peer, and keeps a broadcast's
 * audience on the presenter's slide. Extracted from `PowerPointViewer.svelte`
 * (matching the `usePresentationEffects` / `useViewerEffects` composable
 * pattern already used there) to keep that root component within the
 * file-size budget.
 */
import { createBroadcastFollower } from 'ooxml-ui/pptx';
import { untrack } from 'svelte';

import type { CollaborationController } from './collaboration.svelte';

export interface CollaborationPresenceEffectsDeps {
	collab: CollaborationController;
	getCurrentSlide: () => number;
	getSlideCount: () => number;
	getSelectedElementId: () => string | null;
	goTo: (index: number) => void;
}

export function useCollaborationPresenceEffects(deps: CollaborationPresenceEffectsDeps): void {
	// Publish local active-slide/selection changes.
	$effect(() => {
		if (deps.collab.active) {
			deps.collab.setActiveSlide(deps.getCurrentSlide());
		}
	});
	$effect(() => {
		if (deps.collab.active) {
			deps.collab.setSelection(deps.getSelectedElementId() ?? undefined, deps.getCurrentSlide());
		}
	});
	// Drive follow-mode navigation.
	$effect(() => {
		const followedId = deps.collab.followedClientId;
		if (followedId === null) {
			return;
		}
		const peer = deps.collab.remotePresences.find((p) => p.clientId === followedId);
		if (peer && peer.activeSlideIndex !== deps.getCurrentSlide()) {
			deps.goTo(peer.activeSlideIndex);
		}
	});
	// A broadcast's audience follows the presenter each time the presenter moves
	// (the shared rule: only a local `viewer` follows the `owner`). Following a
	// peer by hand stands this down, so the two never fight over navigation.
	const broadcast = createBroadcastFollower();
	$effect(() => {
		const target = broadcast.step({
			localRole: deps.collab.activeCollaboration?.role,
			peers: deps.collab.remotePresences,
			slideCount: deps.getSlideCount(),
			followingPeer: deps.collab.followedClientId !== null,
			currentSlide: untrack(deps.getCurrentSlide),
		});
		if (target !== null) {
			untrack(() => deps.goTo(target));
		}
	});
}
