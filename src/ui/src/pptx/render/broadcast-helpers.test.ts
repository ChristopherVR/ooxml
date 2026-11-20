/**
 * broadcast-helpers.test.ts: unit tests for the shared Broadcast-dialog
 * helpers (room-id generation, start config, viewer-link building, clipboard
 * detection). Pure helpers, no framework.
 */

import { describe, expect, it } from 'vitest';

import {
	DEFAULT_BROADCAST_SERVER_URL,
	buildBroadcastConfig,
	canStartBroadcast,
	canUseClipboard,
	generateBroadcastRoomId,
	resolveTransportForServerUrl,
	seedBroadcastFields,
} from './broadcast-helpers';
import { buildBroadcastViewerUrl } from './share-session';

describe('generateBroadcastRoomId', () => {
	it('produces a "broadcast-<suffix>" id', () => {
		expect(generateBroadcastRoomId()).toMatch(/^broadcast-[a-z0-9]+$/u);
	});

	it('produces distinct ids across calls', () => {
		const ids = new Set<string>();
		for (let i = 0; i < 50; i++) {
			ids.add(generateBroadcastRoomId());
		}
		expect(ids.size).toBeGreaterThan(1);
	});
});

describe('seedBroadcastFields', () => {
	it('uses the supplied defaults verbatim', () => {
		const fields = seedBroadcastFields({
			roomId: 'broadcast-fixed',
			serverUrl: 'ws://example.test:1234',
		});
		expect(fields).toStrictEqual({
			roomId: 'broadcast-fixed',
			serverUrl: 'ws://example.test:1234',
		});
	});

	it('auto-generates a room id and falls back to the default server', () => {
		const fields = seedBroadcastFields();
		expect(fields.roomId).toMatch(/^broadcast-[a-z0-9]+$/u);
		expect(fields.serverUrl).toBe(DEFAULT_BROADCAST_SERVER_URL);
	});
});

describe('resolveTransportForServerUrl', () => {
	it('maps a blank server URL to webrtc and anything else to websocket', () => {
		expect(resolveTransportForServerUrl('')).toBe('webrtc');
		expect(resolveTransportForServerUrl('   ')).toBe('webrtc');
		expect(resolveTransportForServerUrl('ws://s')).toBe('websocket');
		expect(resolveTransportForServerUrl('wss://collab.example.com')).toBe('websocket');
	});
});

describe('canStartBroadcast', () => {
	it('requires a room id; a blank server URL is allowed (P2P)', () => {
		expect(canStartBroadcast({ roomId: 'r', serverUrl: 'ws://s' })).toBeTruthy();
		expect(canStartBroadcast({ roomId: '', serverUrl: 'ws://s' })).toBeFalsy();
		expect(canStartBroadcast({ roomId: 'r', serverUrl: '  ' })).toBeTruthy();
	});
});

describe('buildBroadcastConfig', () => {
	it('assembles a trimmed config with the derived transport', () => {
		expect(buildBroadcastConfig({ roomId: '  room  ', serverUrl: '  ws://s  ' })).toStrictEqual({
			roomId: 'room',
			serverUrl: 'ws://s',
			transport: 'websocket',
		});
	});

	it('selects the webrtc transport with a fresh secret when the server URL is blank', () => {
		const config = buildBroadcastConfig({ roomId: 'room', serverUrl: '  ' });
		expect(config).toStrictEqual({
			roomId: 'room',
			serverUrl: '',
			transport: 'webrtc',
			secret: expect.stringMatching(/^[A-Za-z0-9_-]{22}$/u),
		});
		expect(buildBroadcastConfig({ roomId: 'room', serverUrl: '' })?.secret).not.toBe(
			config?.secret,
		);
	});

	it('returns null when incomplete', () => {
		expect(buildBroadcastConfig({ roomId: '', serverUrl: 'ws://s' })).toBeNull();
	});
});

describe('buildBroadcastViewerUrl', () => {
	const location = { origin: 'https://app.test', pathname: '/' };

	it('builds a ?broadcast=&server= link from a location', () => {
		const url = buildBroadcastViewerUrl({ roomId: 'room-1', serverUrl: 'ws://x' }, location);
		expect(url).toBe('https://app.test/?broadcast=room-1&server=ws%3A%2F%2Fx');
	});

	it('carries a peer-to-peer broadcast secret in the fragment only', () => {
		const config = buildBroadcastConfig({ roomId: 'room-1', serverUrl: '  ' });
		const url = buildBroadcastViewerUrl(config!, location);
		expect(url).toBe(`https://app.test/?broadcast=room-1&transport=webrtc#key=${config?.secret}`);
	});

	it('falls back to the room id without a location or a peer-to-peer secret', () => {
		expect(buildBroadcastViewerUrl({ roomId: 'room-1', serverUrl: 'ws://x' })).toBe('room-1');
		expect(buildBroadcastViewerUrl({ roomId: 'room-1', serverUrl: '' }, location)).toBe('room-1');
	});
});

describe('canUseClipboard', () => {
	it('is true when navigator.clipboard.writeText exists', () => {
		const nav = { clipboard: { writeText: () => Promise.resolve() } } as unknown as Navigator;
		expect(canUseClipboard(nav)).toBeTruthy();
	});

	it('is false when clipboard or navigator is unavailable', () => {
		expect(canUseClipboard(undefined)).toBeFalsy();
		expect(canUseClipboard({} as Navigator)).toBeFalsy();
	});
});
