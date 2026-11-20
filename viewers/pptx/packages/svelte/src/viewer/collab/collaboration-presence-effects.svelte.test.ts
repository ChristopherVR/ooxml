/**
 * A broadcast's audience follows the presenter: each time the session owner moves to a slide, a
 * local viewer goes there too, and it may still look back at another slide in between.
 */
import { flushSync } from 'svelte';
import { describe, expect, it, vi } from 'vitest';

import { useCollaborationPresenceEffects } from './collaboration-presence-effects.svelte';
import type { CollaborationController } from './collaboration.svelte';

interface Peer {
	clientId: number;
	role?: 'owner' | 'collaborator' | 'viewer';
	activeSlideIndex: number;
}

function audience(role: 'viewer' | 'collaborator') {
	const session = $state({
		active: true,
		followedClientId: null as number | null,
		remotePresences: [] as Peer[],
		activeCollaboration: { role },
		setActiveSlide: () => {},
		setSelection: () => {},
	});
	let current = $state(0);
	const goTo = vi.fn((index: number) => {
		current = index;
	});
	const dispose = $effect.root(() => {
		useCollaborationPresenceEffects({
			collab: session as unknown as CollaborationController,
			getCurrentSlide: () => current,
			getSlideCount: () => 5,
			getSelectedElementId: () => null,
			goTo,
		});
	});
	flushSync();
	const presenterAt = (slide: number) => {
		session.remotePresences = [{ clientId: 7, role: 'owner', activeSlideIndex: slide }];
		flushSync();
	};
	const lookAt = (slide: number) => {
		current = slide;
		flushSync();
	};
	return { goTo, presenterAt, lookAt, dispose, current: () => current };
}

describe('broadcast audience', () => {
	it('follows each move of the presenter', () => {
		const viewer = audience('viewer');
		viewer.presenterAt(2);
		expect(viewer.current()).toBe(2);
		viewer.lookAt(0);
		expect(viewer.current()).toBe(0);
		viewer.presenterAt(3);
		expect(viewer.current()).toBe(3);
		viewer.dispose();
	});

	it('leaves a collaborator free to navigate', () => {
		const peer = audience('collaborator');
		peer.presenterAt(2);
		expect(peer.goTo).not.toHaveBeenCalled();
		peer.dispose();
	});
});
