import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { CollaborationConfig } from '../../types';
import type { CollabProviderHandle } from '../collaboration-provider';
import type { CollabYSession } from '../collaboration-yjs-session';
import { createCollaborationSessionController } from './collaboration-session-controller';

interface FakeProvider extends CollabProviderHandle {
	emitStatus(connected: boolean): void;
	emitSynced(): void;
}

function fakeProvider(connectedNow = false): FakeProvider {
	const status: ((connected: boolean) => void)[] = [];
	const synced: (() => void)[] = [];
	return {
		awareness: {
			clientID: 7,
			setLocalStateField: () => {},
			getStates: () => new Map(),
			on: () => {},
		},
		onStatus: (cb) => status.push(cb),
		onSynced: (cb) => synced.push(cb),
		connectedNow,
		syncedNow: false,
		destroy: vi.fn(),
		emitStatus: (connected) => status.forEach((cb) => cb(connected)),
		emitSynced: () => synced.forEach((cb) => cb()),
	};
}

function fakeSession(provider?: FakeProvider): CollabYSession {
	return {
		doc: {} as CollabYSession['doc'],
		factories: {} as CollabYSession['factories'],
		awareness: provider?.awareness ?? {
			clientID: 9,
			setLocalStateField: () => {},
			getStates: () => new Map(),
			on: () => {},
		},
		...(provider ? { provider } : {}),
		dispose: vi.fn(),
	};
}

const ws: CollaborationConfig = {
	serverUrl: 'wss://collab.example.test',
	roomId: 'room-1',
	userName: 'Ada',
	transport: 'websocket',
};

describe('createCollaborationSessionController', () => {
	beforeEach(() => {
		vi.useFakeTimers();
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('opens a websocket session: connecting, then connected, then synced', async () => {
		const provider = fakeProvider();
		const onOpen = vi.fn();
		const onSynced = vi.fn();
		const controller = createCollaborationSessionController({
			openSession: async () => fakeSession(provider),
			onOpen,
			onSynced,
			registerTeardown: false,
		});
		const started = controller.start(ws);
		expect(controller.getSnapshot().status).toBe('connecting');
		await started;
		expect(onOpen).toHaveBeenCalledOnce();
		expect(controller.getSnapshot()).toMatchObject({ active: true, clientId: 7 });
		provider.emitStatus(true);
		expect(controller.getSnapshot().status).toBe('connected');
		provider.emitSynced();
		expect(onSynced).toHaveBeenCalledOnce();
		expect(controller.getSnapshot().synced).toBe(true);
	});

	it('errors when the websocket never connects', async () => {
		const provider = fakeProvider();
		const controller = createCollaborationSessionController({
			openSession: async () => fakeSession(provider),
			connectTimeoutMs: 1000,
			registerTeardown: false,
		});
		await controller.start(ws);
		await vi.advanceTimersByTimeAsync(1000);
		expect(controller.getSnapshot()).toMatchObject({ status: 'error', active: false });
		expect(provider.destroy).not.toHaveBeenCalled();
	});

	it('is connected at once over webrtc, and a drop re-gates writes', async () => {
		const provider = fakeProvider();
		const controller = createCollaborationSessionController({
			openSession: async () => fakeSession(provider),
			registerTeardown: false,
		});
		await controller.start({ ...ws, transport: 'webrtc' });
		expect(controller.getSnapshot().status).toBe('connected');
		controller.gate.open();
		provider.emitStatus(false);
		expect(controller.getSnapshot()).toMatchObject({ status: 'disconnected', synced: false });
		expect(controller.gate.isOpen()).toBe(false);
	});

	it('refuses an invalid room and a ws:// socket from a secure page', async () => {
		const openSession = vi.fn(async () => fakeSession(fakeProvider()));
		const controller = createCollaborationSessionController({
			openSession,
			registerTeardown: false,
		});
		await controller.start({ ...ws, roomId: '' });
		expect(controller.getSnapshot().status).toBe('error');
		vi.stubGlobal('window', { location: { protocol: 'https:' } });
		try {
			await controller.start({ ...ws, serverUrl: 'ws://insecure.example.test' });
			expect(controller.getSnapshot().status).toBe('error');
		} finally {
			vi.unstubAllGlobals();
		}
		expect(openSession).not.toHaveBeenCalled();
	});

	it('stop closes the binding layer, disposes the session and resets', async () => {
		const provider = fakeProvider(true);
		const session = fakeSession(provider);
		const onClose = vi.fn();
		const controller = createCollaborationSessionController({
			openSession: async () => session,
			onClose,
			registerTeardown: false,
		});
		await controller.start(ws);
		expect(controller.getSnapshot().status).toBe('connected');
		onClose.mockClear();
		controller.stop();
		expect(onClose).toHaveBeenCalledOnce();
		expect(session.dispose).toHaveBeenCalledOnce();
		expect(controller.getSnapshot()).toMatchObject({
			status: 'disconnected',
			active: false,
			config: null,
		});
	});

	it('drops a session that resolves after a newer start', async () => {
		const first = fakeSession(fakeProvider());
		let resolveFirst: ((value: CollabYSession) => void) | undefined;
		const openSession = vi
			.fn()
			.mockImplementationOnce(
				() =>
					new Promise<CollabYSession>((resolve) => {
						resolveFirst = resolve;
					}),
			)
			.mockResolvedValueOnce(fakeSession(fakeProvider(true)));
		const controller = createCollaborationSessionController({
			openSession,
			registerTeardown: false,
		});
		const pending = controller.start(ws);
		await controller.start({ ...ws, roomId: 'room-2' });
		resolveFirst?.(first);
		await pending;
		expect(first.dispose).toHaveBeenCalledOnce();
		expect(controller.getSnapshot().config?.roomId).toBe('room-2');
	});

	it('leaves the status of a borrowed host session to the binding', async () => {
		const controller = createCollaborationSessionController({
			openSession: async () => fakeSession(),
			registerTeardown: false,
		});
		await controller.start({
			...ws,
			externalSession: {} as NonNullable<CollaborationConfig['externalSession']>,
		});
		expect(controller.getSnapshot()).toMatchObject({ status: 'connecting', active: true });
		controller.setStatus('connected');
		expect(controller.getSnapshot().status).toBe('connected');
	});
});
