/**
 * collaboration-broadcast-follow.ts: the single, framework-agnostic rule for
 * whether the local user should auto-follow a broadcaster's active slide.
 *
 * A "broadcast" session has one presenter that joins with `role: 'owner'`.
 * Auto-follow is one-way: only a local `viewer` is pulled along to the
 * owner's slide. A `collaborator` (two-way editing peer) and the `owner`
 * itself navigate freely and are never yanked.
 *
 * Every binding (React/Vue/Angular/Svelte/Vanilla) must consult this helper
 * instead of hand-rolling the guard, so the policy cannot drift per binding
 * (previously React/Angular gated on `localRole === 'viewer'` while Vue
 * followed the owner regardless of the local role).
 */

import type { CollaborationRole } from '../types';

export interface ShouldAutoFollowBroadcasterInput {
	/** The local user's role in the session (undefined when not in a session). */
	localRole: CollaborationRole | undefined;
	/**
	 * The role of the peer being considered as the broadcaster. Auto-follow
	 * only applies to the session `owner`; pass the candidate peer's role.
	 */
	broadcasterRole: CollaborationRole | undefined;
}

/**
 * True when the local user should auto-follow the given broadcaster's active
 * slide: the local user is a `viewer` and the broadcaster is the `owner`.
 */
export function shouldAutoFollowBroadcaster(input: ShouldAutoFollowBroadcasterInput): boolean {
	return input.localRole === 'viewer' && input.broadcasterRole === 'owner';
}

export interface BroadcastFollowSlideInput {
	/** The local user's role in the session (undefined when not in a session). */
	localRole: CollaborationRole | undefined;
	/** The remote peers' presence, as every binding's presence list holds it. */
	peers: readonly { role?: CollaborationRole | undefined; activeSlideIndex: number }[];
	/** How many slides the local deck has. */
	slideCount: number;
}

/**
 * The slide a local `viewer` shows to follow the session owner, or null when it
 * follows no one (not a viewer, no owner present, or the owner is on a slide the
 * local deck does not have yet). A binding navigates when this value CHANGES, so
 * the audience can still look back at another slide until the presenter moves.
 */
export function broadcastFollowSlide(input: BroadcastFollowSlideInput): number | null {
	const owner = input.peers.find((peer) => peer.role === 'owner');
	if (
		!owner ||
		!shouldAutoFollowBroadcaster({ localRole: input.localRole, broadcasterRole: owner.role })
	) {
		return null;
	}
	const slide = owner.activeSlideIndex;
	return slide >= 0 && slide < input.slideCount ? slide : null;
}

export interface BroadcastFollowStepInput extends BroadcastFollowSlideInput {
	/** The local user follows a peer by hand (the follow bar), which stands auto-follow down. */
	followingPeer: boolean;
	/** The slide the local user is on. */
	currentSlide: number;
}

/**
 * A broadcast audience's navigation for bindings that react to presence
 * changes (Svelte and vanilla): `step` returns the slide to go to when the
 * presenter has moved since the last step, or null to stay where the local
 * user is.
 */
export function createBroadcastFollower(): {
	step(input: BroadcastFollowStepInput): number | null;
} {
	let lastTarget: number | null = null;
	return {
		step(input) {
			const target = broadcastFollowSlide(input);
			if (target === null || input.followingPeer || target === lastTarget) {
				lastTarget = input.followingPeer ? null : target;
				return null;
			}
			lastTarget = target;
			return target === input.currentSlide ? null : target;
		},
	};
}
