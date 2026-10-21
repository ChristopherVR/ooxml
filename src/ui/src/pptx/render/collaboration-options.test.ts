import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
import { describe, expect, it } from 'vitest';

import type { CollaborationConfig } from '../types';
import {
	externalSessionFromCollabSession,
	resolveCollaborationConfig,
} from './collaboration-options';

function makeSession() {
	const hub = createMemoryHub();
	return createCollabSession({
		roomId: 'room-1',
		provider: transportProvider({ transport: hub.createTransport('room-1') }),
		user: { name: 'Ada' },
		heartbeatMs: 0,
	});
}

describe('resolveCollaborationConfig', () => {
	it('returns undefined for no input and the same object for a legacy config', () => {
		expect(resolveCollaborationConfig(undefined)).toBeUndefined();
		const legacy: CollaborationConfig = { roomId: 'r', serverUrl: 'ws://x', userName: 'A' };
		expect(resolveCollaborationConfig(legacy)).toBe(legacy);
	});

	it('maps a host session to a stable borrowed externalSession', () => {
		const session = makeSession();
		const input = {
			session,
			user: { name: 'Bo', color: '#112233', role: 'viewer' as const },
			intent: 'join' as const,
		};
		const config = resolveCollaborationConfig(input)!;
		expect(config).toMatchObject({
			roomId: 'room-1',
			serverUrl: '',
			userName: 'Bo',
			userColor: '#112233',
			role: 'viewer',
			sessionIntent: 'join',
		});
		expect(config.externalSession).toBe(externalSessionFromCollabSession(session));
		expect(resolveCollaborationConfig(input)).toBe(config);
		expect(resolveCollaborationConfig({ session })!.externalSession).toBe(config.externalSession);
		expect(config.externalSession!.getSnapshot()).toBe(config.externalSession!.getSnapshot());
		session.destroy();
	});

	it('lets explicit legacy fields win over the shared ones', () => {
		const session = makeSession();
		const config = resolveCollaborationConfig({
			session,
			userName: 'Legacy',
			user: { name: 'Shared' },
			sessionIntent: 'create',
			intent: 'join',
		} as never)!;
		expect(config.userName).toBe('Legacy');
		expect(config.sessionIntent).toBe('create');
		session.destroy();
	});

	it('defaults the user name to the session identity', () => {
		const session = makeSession();
		expect(resolveCollaborationConfig({ session })!.userName).toBe('Ada');
		session.destroy();
	});

	it('resolves a provider without a session to an invalid room so bindings report an error', () => {
		const config = resolveCollaborationConfig({
			roomId: 'ok',
			provider: (() => {}) as never,
		} as never)!;
		expect(config.roomId).toBe('');
	});
});
