import { describe, expect, it } from 'vitest';

import { translationsEn } from '../i18n/translations-en';
import type { CollaborationFailure } from './collaboration-failure';
import {
	CollaborationStartError,
	collaborationFailureMessageKey,
	collaborationFailureOf,
	loadCollaborationModule,
} from './collaboration-failure';

const FAILURES: CollaborationFailure[] = [
	'invalid-room',
	'missing-secret',
	'insecure-page',
	'module-unavailable',
	'insecure-server',
	'server-unreachable',
	'no-broadcast',
	'unknown',
];

describe('collaborationFailureMessageKey', () => {
	it('gives every reason its own English sentence while the session is in error', () => {
		const texts = FAILURES.map((failure) => {
			const key = collaborationFailureMessageKey('error', failure);
			return key ? translationsEn[key] : undefined;
		});
		expect(texts.every((text) => typeof text === 'string' && text.length > 0)).toBe(true);
		expect(new Set(texts).size).toBe(FAILURES.length);
	});

	it('blames no server for a peer-to-peer failure', () => {
		for (const failure of ['missing-secret', 'insecure-page', 'unknown'] as const) {
			const key = collaborationFailureMessageKey('error', failure)!;
			expect(translationsEn[key]).not.toMatch(/server/iu);
		}
	});

	it('explains an error without a reason neutrally, and says nothing outside an error', () => {
		expect(collaborationFailureMessageKey('error', null)).toBe(
			'pptx.collaboration.failure.unknown',
		);
		for (const status of ['connected', 'connecting', 'disconnected'] as const) {
			expect(collaborationFailureMessageKey(status, 'missing-secret')).toBeNull();
		}
	});
});

describe('collaborationFailureOf', () => {
	it('reads the reason a start error carries, and anything else as unknown', () => {
		expect(collaborationFailureOf(new CollaborationStartError('no-broadcast', 'x'))).toBe(
			'no-broadcast',
		);
		expect(collaborationFailureOf({ failure: 'missing-secret' })).toBe('missing-secret');
		expect(collaborationFailureOf({ failure: 'made-up' })).toBe('unknown');
		expect(collaborationFailureOf(new Error('boom'))).toBe('unknown');
		expect(collaborationFailureOf(null)).toBe('unknown');
	});

	it('classifies a rejected module import as module-unavailable', async () => {
		const failed = await loadCollaborationModule('y-webrtc', () =>
			Promise.reject(new Error('Failed to fetch dynamically imported module')),
		).catch((error: unknown) => error);
		expect(collaborationFailureOf(failed)).toBe('module-unavailable');
		expect(await loadCollaborationModule('yjs', async () => 42)).toBe(42);
	});
});
