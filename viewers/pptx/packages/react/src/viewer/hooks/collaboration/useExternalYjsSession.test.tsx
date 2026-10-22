import { createCollabSession, createMemoryHub, transportProvider } from 'ooxml-core/collab';
// @vitest-environment happy-dom
import type {
	CollaborationConfig,
	ExternalCollaborationSession,
	ExternalCollaborationSnapshot,
} from 'ooxml-ui/pptx';
import { resolveCollaborationConfig } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Awareness } from 'y-protocols/awareness';
import { Doc } from 'yjs';

import type { CollaborationContextValue } from './types';
import { useCollaborativeState } from './useCollaborativeState';

const transports = vi.hoisted(() => ({ websocket: vi.fn(), webrtc: vi.fn() }));
vi.mock(import('y-websocket'), () => ({ WebsocketProvider: transports.websocket }));
vi.mock(import('y-webrtc'), () => ({ WebrtcProvider: transports.webrtc }));

let root: Root;
let container: HTMLDivElement;
let current: CollaborationContextValue | null;
const resources: { doc: Doc; awareness: Awareness }[] = [];

function createSession() {
	const doc = new Doc();
	const awareness = new Awareness(doc);
	awareness.setLocalState({ hostField: 'keep', presence: { hostPresence: true } });
	resources.push({ doc, awareness });
	const listeners = new Set<() => void>();
	let snapshot: ExternalCollaborationSnapshot = { status: 'connected', synced: false };
	const session: ExternalCollaborationSession = {
		doc,
		awareness,
		getSnapshot: () => snapshot,
		subscribe: (listener) => {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
	};
	return {
		doc,
		awareness,
		listeners,
		session,
		update(next: ExternalCollaborationSnapshot) {
			snapshot = next;
			listeners.forEach((listener) => listener());
		},
	};
}

function Probe({ config }: { config?: CollaborationConfig }) {
	current = useCollaborativeState({ config, canvasWidth: 960, canvasHeight: 540 });
	return null;
}

function render(session: ExternalCollaborationSession, userName = 'Alice') {
	act(() => {
		root.render(
			<Probe config={{ roomId: 'external', serverUrl: '', userName, externalSession: session }} />,
		);
	});
}

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	current = null;
	vi.clearAllMocks();
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	for (const { doc, awareness } of resources.splice(0)) {
		awareness.destroy();
		doc.destroy();
	}
	vi.useRealTimers();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('host-owned collaboration sessions', () => {
	it('stays attached after a cancellable beforeunload and leaves on pagehide', () => {
		const host = createSession();
		render(host.session);
		act(() => window.dispatchEvent(new Event('beforeunload', { cancelable: true })));
		expect(current?.doc).toBe(host.doc);
		expect(host.listeners.size).toBe(1);
		act(() => window.dispatchEvent(new Event('pagehide')));
		expect(host.listeners.size).toBe(0);
		expect(host.doc.isDestroyed).toBeFalsy();
	});

	it('borrows a shared-contract CollabSession passed as collaboration.session', () => {
		const hub = createMemoryHub();
		const shared = createCollabSession({
			roomId: 'shared-room',
			provider: transportProvider({ transport: hub.createTransport('shared-room') }),
			user: { name: 'Ada' },
			heartbeatMs: 0,
		});
		const input = { session: shared, user: { name: 'Bo' } };
		const config = resolveCollaborationConfig(input);
		act(() => root.render(<Probe config={config} />));
		expect(current?.doc).toBe(shared.doc);
		expect(current?.config.roomId).toBe('shared-room');
		expect(current?.config.userName).toBe('Bo');
		act(() => root.render(<Probe config={resolveCollaborationConfig(input)} />));
		expect(current?.doc).toBe(shared.doc);
		act(() => root.render(null));
		expect(shared.doc.isDestroyed).toBeFalsy();
		shared.destroy();
	});

	it('retains one borrowed subscription through StrictMode effect replay', () => {
		const host = createSession();
		act(() =>
			root.render(
				<React.StrictMode>
					<Probe
						config={{
							roomId: 'strict',
							serverUrl: '',
							userName: 'Alice',
							externalSession: host.session,
						}}
					/>
				</React.StrictMode>,
			),
		);
		expect(current?.doc).toBe(host.doc);
		expect(host.listeners.size).toBe(1);
		expect(host.awareness.getLocalState()?.presence.userName).toBe('Alice');
		act(() => root.render(null));
		expect(host.listeners.size).toBe(0);
		expect(host.awareness.getLocalState()).toStrictEqual({
			hostField: 'keep',
			presence: { hostPresence: true },
		});
		expect(host.doc.isDestroyed).toBeFalsy();
	});

	it('uses the host document and waits for its sync signal without creating a transport', () => {
		vi.useFakeTimers();
		const host = createSession();
		render(host.session);
		expect(current?.doc).toBe(host.doc);
		expect(current?.status).toBe('connected');
		act(() => vi.advanceTimersByTime(60_000));
		expect(current?.synced).toBeFalsy();
		act(() => host.update({ status: 'connected', synced: true }));
		expect(current?.synced).toBeTruthy();
		expect(transports.websocket).not.toHaveBeenCalled();
		expect(transports.webrtc).not.toHaveBeenCalled();
	});

	it('preserves host resources and unrelated awareness on detach', () => {
		const host = createSession();
		const destroyDoc = vi.spyOn(host.doc, 'destroy');
		const destroyAwareness = vi.spyOn(host.awareness, 'destroy');
		render(host.session);
		expect(host.awareness.getLocalState()?.presence.userName).toBe('Alice');
		expect(host.listeners.size).toBe(1);
		act(() => root.render(<Probe />));
		expect(host.listeners.size).toBe(0);
		expect(host.awareness.getLocalState()).toStrictEqual({
			hostField: 'keep',
			presence: { hostPresence: true },
		});
		expect(destroyDoc).not.toHaveBeenCalled();
		expect(destroyAwareness).not.toHaveBeenCalled();
		host.doc.getMap('host').set('stillAlive', true);
		expect(host.doc.getMap('host').get('stillAlive')).toBeTruthy();
	});

	it('updates connection state and identity without replacing host resources', () => {
		const host = createSession();
		render(host.session);
		render(host.session, 'Renamed');
		expect(host.listeners.size).toBe(1);
		expect(host.awareness.getLocalState()?.presence.userName).toBe('Renamed');
		act(() => host.update({ status: 'disconnected', synced: true }));
		expect(current?.doc).toBe(host.doc);
		expect(current?.status).toBe('disconnected');
		expect(current?.synced).toBeTruthy();
		act(() => host.update({ status: 'connecting', synced: false }));
		expect(current?.synced).toBeFalsy();
	});

	it('unsubscribes from a replaced session and ignores its delayed events', () => {
		const first = createSession();
		const second = createSession();
		render(first.session);
		const staleListener = [...first.listeners][0];
		render(second.session);
		expect(first.listeners.size).toBe(0);
		expect(second.listeners.size).toBe(1);
		act(() => staleListener());
		expect(current?.doc).toBe(second.doc);
		expect(first.awareness.getLocalState()?.presence).toStrictEqual({ hostPresence: true });
	});

	it('detaches on pagehide and reattaches on a persisted pageshow without destroying the host', () => {
		const host = createSession();
		const destroy = vi.spyOn(host.doc, 'destroy');
		render(host.session);
		act(() => window.dispatchEvent(new Event('pagehide')));
		expect(host.listeners.size).toBe(0);
		expect(host.awareness.getLocalState()?.hostField).toBe('keep');
		const event = new Event('pageshow');
		Object.defineProperty(event, 'persisted', { value: true });
		act(() => window.dispatchEvent(event));
		expect(current?.doc).toBe(host.doc);
		expect(host.listeners.size).toBe(1);
		expect(destroy).not.toHaveBeenCalled();
	});

	it('releases its subscription when the host snapshot getter fails', () => {
		const host = createSession();
		host.session.getSnapshot = () => {
			throw new Error('session unavailable');
		};
		render(host.session);
		expect(current?.status).toBe('error');
		expect(host.listeners.size).toBe(0);
		expect(host.awareness.getLocalState()?.presence).toStrictEqual({ hostPresence: true });
	});
});

describe('viewer-owned collaboration sessions', () => {
	const settle = () =>
		act(async () => void (await new Promise((resolve) => setTimeout(resolve, 20))));

	it('connects a provider-only config to the provider room and destroys it on stop', async () => {
		const hub = createMemoryHub();
		const host = createCollabSession({
			roomId: 'owned-room',
			provider: transportProvider({ transport: hub.createTransport('owned-room') }),
			user: { name: 'Host' },
			heartbeatMs: 0,
		});
		const input = {
			roomId: 'owned-room',
			provider: transportProvider({ transport: hub.createTransport('owned-room') }),
			user: { name: 'Ada' },
			intent: 'join' as const,
		};
		act(() => root.render(<Probe config={resolveCollaborationConfig(input)} />));
		expect(current?.status).toBe('connecting');
		await settle();
		expect(current?.config.externalSession).toBeDefined();
		expect(current?.config.roomId).toBe('owned-room');
		expect(current?.status).toBe('connected');
		host.doc.getMap('shared').set('k', 'v');
		expect(current?.doc?.getMap('shared').get('k')).toBe('v');
		const awareness = current!.config.externalSession!.awareness;
		act(() => root.render(<Probe />));
		expect(awareness.getLocalState()).toBeNull();
		expect(host.awareness.getLocalState()).not.toBeNull();
		host.destroy();
	});

	it('shares between windows of this browser without any server', async () => {
		act(() =>
			root.render(
				<Probe config={resolveCollaborationConfig({ roomId: 'tabs', user: { name: 'Tab' } })} />,
			),
		);
		await settle();
		expect(current?.config.externalSession).toBeDefined();
		expect(current?.status).toBe('connected');
	});

	it('reports an error for an invalid room and opens nothing when stopped while pending', async () => {
		act(() =>
			root.render(
				<Probe config={resolveCollaborationConfig({ roomId: '', user: { name: 'A' } })} />,
			),
		);
		await settle();
		expect(current?.status).toBe('error');
		act(() =>
			root.render(
				<Probe config={resolveCollaborationConfig({ roomId: 'pending', user: { name: 'A' } })} />,
			),
		);
		act(() => root.render(<Probe />));
		await settle();
		expect(current).toBeNull();
	});
});
