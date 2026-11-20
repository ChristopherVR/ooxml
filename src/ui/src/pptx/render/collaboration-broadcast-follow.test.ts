import { describe, expect, it } from 'vitest';

import {
	broadcastFollowSlide,
	shouldAutoFollowBroadcaster,
} from './collaboration-broadcast-follow';

describe('shouldAutoFollowBroadcaster', () => {
	it('a viewer auto-follows the owner broadcaster', () => {
		expect(
			shouldAutoFollowBroadcaster({ localRole: 'viewer', broadcasterRole: 'owner' }),
		).toBeTruthy();
	});

	it('a collaborator does NOT auto-follow the owner', () => {
		expect(
			shouldAutoFollowBroadcaster({ localRole: 'collaborator', broadcasterRole: 'owner' }),
		).toBeFalsy();
	});

	it('the owner does NOT auto-follow itself', () => {
		expect(
			shouldAutoFollowBroadcaster({ localRole: 'owner', broadcasterRole: 'owner' }),
		).toBeFalsy();
	});

	it('a viewer does NOT follow a non-owner (collaborator) peer', () => {
		expect(
			shouldAutoFollowBroadcaster({ localRole: 'viewer', broadcasterRole: 'collaborator' }),
		).toBeFalsy();
	});

	it('an undefined local role never auto-follows', () => {
		expect(
			shouldAutoFollowBroadcaster({ localRole: undefined, broadcasterRole: 'owner' }),
		).toBeFalsy();
	});
});

describe('broadcastFollowSlide', () => {
	const peers = [
		{ role: 'viewer' as const, activeSlideIndex: 0 },
		{ role: 'owner' as const, activeSlideIndex: 2 },
	];

	it('is the owner slide for a local viewer', () => {
		expect(broadcastFollowSlide({ localRole: 'viewer', peers, slideCount: 5 })).toBe(2);
	});

	it('is null for anyone but a viewer, and without an owner', () => {
		expect(broadcastFollowSlide({ localRole: 'collaborator', peers, slideCount: 5 })).toBeNull();
		expect(broadcastFollowSlide({ localRole: undefined, peers, slideCount: 5 })).toBeNull();
		expect(
			broadcastFollowSlide({ localRole: 'viewer', peers: peers.slice(0, 1), slideCount: 5 }),
		).toBeNull();
	});

	it('is null while the local deck does not have the owner slide yet', () => {
		expect(broadcastFollowSlide({ localRole: 'viewer', peers, slideCount: 1 })).toBeNull();
	});
});
