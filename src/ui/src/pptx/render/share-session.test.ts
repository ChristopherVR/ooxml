import { describe, expect, it } from 'vitest';

import { webrtcRoomName } from 'ooxml-core/collab';

import {
	buildCollaborationShareUrl,
	buildCreateCollaborationConfig,
	buildJoinCollaborationConfig,
	shareFormFailureKey,
} from './share-session';

const SECRET = 'AAECAwQFBgcICQoLDA0ODw';

describe('share session helpers', () => {
	it('cannot join a peer-to-peer session from a bare room code', () => {
		const config = { roomId: 'p2p-room', serverUrl: '', transport: 'webrtc' as const };
		const invitation = buildCollaborationShareUrl(
			{ ...config, secret: SECRET },
			{ origin: 'null', pathname: '/' },
		);
		expect(invitation).toBe(config.roomId);
		expect(
			buildJoinCollaborationConfig({ invitation, userName: 'Guest', serverUrl: '' }),
		).toBeNull();
	});
	it('creates a peer-to-peer session with a fresh secret whose link round-trips', () => {
		const created = buildCreateCollaborationConfig({
			roomId: 'quarterly-board-deck',
			userName: 'Ada',
			serverUrl: '  ',
		});
		expect(created).toMatchObject({ transport: 'webrtc', sessionIntent: 'create' });
		expect(created?.secret).toMatch(/^[A-Za-z0-9_-]{22}$/u);
		const again = buildCreateCollaborationConfig({
			roomId: 'quarterly-board-deck',
			userName: 'Ada',
			serverUrl: '',
		});
		expect(again?.secret).not.toBe(created?.secret);
		const link = buildCollaborationShareUrl(created!, {
			origin: 'https://app.test',
			pathname: '/viewer',
		});
		const url = new URL(link);
		expect(url.search).toBe('?room=quarterly-board-deck&transport=webrtc');
		expect(url.hash).toBe(`#key=${created?.secret}`);
		expect(
			buildJoinCollaborationConfig({ invitation: link, userName: 'Bo', serverUrl: '' }),
		).toMatchObject({ roomId: 'quarterly-board-deck', secret: created?.secret });
		expect(webrtcRoomName(created!.secret!)).not.toContain('quarterly');
	});
	it('refuses a peer-to-peer link without its secret', () => {
		expect(
			buildJoinCollaborationConfig({
				invitation: 'https://app.test/?room=deck-1&transport=webrtc',
				userName: 'Bo',
				serverUrl: 'wss://fallback',
			}),
		).toBeNull();
		expect(
			buildCollaborationShareUrl(
				{ roomId: 'deck-1', serverUrl: '', transport: 'webrtc' },
				{ origin: 'https://app.test', pathname: '/' },
			),
		).toBe('deck-1');
	});
	it('retains the absent-location room code and exposes malformed nonopaque locations', () => {
		const config = { roomId: 'external-room', serverUrl: 'wss://relay.example' };
		expect(buildCollaborationShareUrl(config)).toBe(config.roomId);
		expect(() =>
			buildCollaborationShareUrl(config, { origin: 'not a URL', pathname: '/' }),
		).toThrow();
	});
	it.each(['null', ''])('keeps a joinable room code for unavailable origin %s', (origin) => {
		const config = { roomId: 'external-room', serverUrl: 'wss://relay.example' };
		const invitation = buildCollaborationShareUrl(config, { origin, pathname: '/' });
		expect(invitation).toBe(config.roomId);
		expect(
			buildJoinCollaborationConfig({ invitation, userName: 'Host', serverUrl: config.serverUrl }),
		).toMatchObject({ roomId: config.roomId, serverUrl: config.serverUrl });
	});
	it('builds an explicitly created websocket session', () => {
		expect(
			buildCreateCollaborationConfig({
				roomId: 'team-room',
				userName: ' Ada ',
				serverUrl: ' wss://x ',
			}),
		).toMatchObject({
			roomId: 'team-room',
			userName: 'Ada',
			serverUrl: 'wss://x',
			transport: 'websocket',
			sessionIntent: 'create',
		});
	});

	it('joins a URL created by another binding and preserves WebRTC signaling', () => {
		const config = buildJoinCollaborationConfig({
			invitation: `https://react.example/viewer?room=deck-1&transport=webrtc&signaling=wss%3A%2F%2Fa%2Cwss%3A%2F%2Fb#key=${SECRET}`,
			userName: 'Grace',
			serverUrl: 'wss://fallback',
		});
		expect(config).toStrictEqual({
			roomId: 'deck-1',
			userName: 'Grace',
			serverUrl: '',
			transport: 'webrtc',
			signaling: ['wss://a', 'wss://b'],
			secret: SECRET,
			role: 'collaborator',
			sessionIntent: 'join',
		});
	});

	it('joins a bare room with the local default server', () => {
		expect(
			buildJoinCollaborationConfig({
				invitation: 'deck_2',
				userName: 'Lin',
				serverUrl: 'wss://vue.example',
			}),
		).toMatchObject({ roomId: 'deck_2', serverUrl: 'wss://vue.example', sessionIntent: 'join' });
	});

	it('rejects invalid or broadcast invitations', () => {
		expect(
			buildJoinCollaborationConfig({ invitation: 'bad room', userName: 'Lin', serverUrl: '' }),
		).toBeNull();
		expect(
			buildJoinCollaborationConfig({
				invitation: 'https://x.test/?broadcast=room-1',
				userName: 'Lin',
				serverUrl: '',
			}),
		).toBeNull();
	});

	it('round-trips a websocket invitation independently of its host app', () => {
		const link = buildCollaborationShareUrl(
			{ roomId: 'deck-3', serverUrl: 'wss://relay.example', transport: 'websocket' },
			{ origin: 'https://angular.example', pathname: '/viewer' },
		);
		expect(
			buildJoinCollaborationConfig({ invitation: link, userName: 'Sam', serverUrl: '' }),
		).toMatchObject({ roomId: 'deck-3', serverUrl: 'wss://relay.example', sessionIntent: 'join' });
	});

	it('names why a keyless peer-to-peer invitation cannot join', () => {
		const keyless = 'https://x.test/?room=deck-4&transport=webrtc';
		const missingSecret = 'pptx.collaboration.failure.missingSecret';
		for (const invitation of [keyless, `${keyless}&key=${SECRET}`]) {
			expect(
				buildJoinCollaborationConfig({ invitation, userName: 'Sam', serverUrl: '' }),
			).toBeNull();
			expect(shareFormFailureKey({ mode: 'join', invitation }, null)).toBe(missingSecret);
			// The invitation's reason replaces an older session failure on the join tab only.
			expect(
				shareFormFailureKey(
					{ mode: 'join', invitation },
					'pptx.collaboration.failure.serverUnreachable',
				),
			).toBe(missingSecret);
			expect(shareFormFailureKey({ mode: 'create', invitation }, null)).toBeNull();
		}
		const complete = `${keyless}#key=${SECRET}`;
		expect(shareFormFailureKey({ mode: 'join', invitation: complete }, null)).toBeNull();
		expect(shareFormFailureKey({ mode: 'join', invitation: 'deck-4' }, 'session-key')).toBe(
			'session-key',
		);
	});
});
