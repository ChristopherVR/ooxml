/**
 * share-helpers.test.ts: Unit tests for the Share-dialog helpers. Ports the
 * Vue `ShareDialog.test.ts` coverage (prefill, config assembly, validation)
 * plus the React share-link builder, against the pure helpers, no TestBed.
 */

import type { CollaborationConfig } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { buildCollaborationConfig, canStartShare, seedShareFields } from './share-helpers';

describe('seedShareFields', () => {
	it('prefills the fields from the defaults', () => {
		const fields = seedShareFields({
			roomId: 'team-room',
			userName: 'Ada',
			serverUrl: 'wss://collab.example.com',
		});
		expect(fields).toStrictEqual({
			roomId: 'team-room',
			userName: 'Ada',
			serverUrl: 'wss://collab.example.com',
		});
	});

	it('coerces absent defaults to empty strings', () => {
		expect(seedShareFields()).toStrictEqual({ roomId: '', userName: '', serverUrl: '' });
		expect(seedShareFields({ roomId: 'room' })).toStrictEqual({
			roomId: 'room',
			userName: '',
			serverUrl: '',
		});
	});
});

describe('canStartShare', () => {
	it('is true when room id + name are non-blank (server optional)', () => {
		expect(canStartShare({ roomId: 'r', userName: 'u', serverUrl: 'wss://s' })).toBeTruthy();
		// Blank server URL is allowed: it selects the serverless P2P transport.
		expect(canStartShare({ roomId: 'r', userName: 'u', serverUrl: '' })).toBeTruthy();
	});

	it('is false when room id or name is blank or whitespace-only', () => {
		expect(canStartShare({ roomId: '', userName: 'u', serverUrl: 'wss://s' })).toBeFalsy();
		expect(canStartShare({ roomId: 'r', userName: '   ', serverUrl: 'wss://s' })).toBeFalsy();
	});
});

describe('buildCollaborationConfig', () => {
	it('assembles a trimmed websocket CollaborationConfig from the fields', () => {
		const config = buildCollaborationConfig({
			roomId: '  edited-room  ',
			userName: 'Grace',
			serverUrl: 'wss://collab.example.com',
		});
		const expected: CollaborationConfig = {
			roomId: 'edited-room',
			userName: 'Grace',
			serverUrl: 'wss://collab.example.com',
			transport: 'websocket',
		};
		expect(config).toStrictEqual(expected);
	});

	it('selects the webrtc transport with a fresh secret when the server URL is blank', () => {
		const config = buildCollaborationConfig({ roomId: 'room', userName: 'Ada', serverUrl: '  ' });
		expect(config).toStrictEqual({
			roomId: 'room',
			userName: 'Ada',
			serverUrl: '',
			transport: 'webrtc',
			secret: expect.stringMatching(/^[A-Za-z0-9_-]{22}$/u),
		});
	});

	it('returns null when required fields are blank', () => {
		expect(buildCollaborationConfig({ roomId: '', userName: '', serverUrl: '' })).toBeNull();
	});
});
