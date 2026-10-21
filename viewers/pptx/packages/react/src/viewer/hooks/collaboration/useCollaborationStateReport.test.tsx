// @vitest-environment happy-dom
import type { CollaborationState } from 'ooxml-core/collab';
import type {
	CollaborationConfig,
	ExternalCollaborationSession,
	ExternalCollaborationSnapshot,
} from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Awareness } from 'y-protocols/awareness';
import { Doc } from 'yjs';

import { useCollaborationStateReport } from './useCollaborationStateReport';
import { useCollaborativeState } from './useCollaborativeState';

vi.mock(import('y-websocket'), () => ({ WebsocketProvider: vi.fn() }));
vi.mock(import('y-webrtc'), () => ({ WebrtcProvider: vi.fn() }));

let root: Root;
let container: HTMLDivElement;
const resources: { doc: Doc; awareness: Awareness }[] = [];

function createSession() {
	const doc = new Doc();
	const awareness = new Awareness(doc);
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
		awareness,
		session,
		update(next: ExternalCollaborationSnapshot) {
			snapshot = next;
			listeners.forEach((listener) => listener());
		},
	};
}

function Probe({
	config,
	onChange,
}: {
	config?: CollaborationConfig;
	onChange: (state: CollaborationState) => void;
}) {
	const collab = useCollaborativeState({ config, canvasWidth: 960, canvasHeight: 540 });
	useCollaborationStateReport(collab, onChange);
	return null;
}

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	for (const { doc, awareness } of resources.splice(0)) {
		awareness.destroy();
		doc.destroy();
	}
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('onCollaborationChange reporting', () => {
	it('reports a host-owned session with its people, then the off state on stop', () => {
		const host = createSession();
		const onChange = vi.fn<(state: CollaborationState) => void>();
		const config: CollaborationConfig = {
			roomId: 'host-room',
			serverUrl: '',
			userName: 'Alice',
			externalSession: host.session,
		};
		act(() => root.render(<Probe config={config} onChange={onChange} />));
		host.update({ status: 'connected', synced: true });
		act(() => {
			host.awareness.states.set(99, {
				presence: {
					clientId: 99,
					userName: 'Bo',
					userColor: '#ef4444',
					activeSlideIndex: 0,
					cursorX: 1,
					cursorY: 1,
					lastUpdated: new Date().toISOString(),
				},
			});
			host.awareness.emit('change', [{ added: [99], updated: [], removed: [] }, 'remote']);
		});
		const last = onChange.mock.calls.at(-1)?.[0];
		expect(last).toMatchObject({
			active: true,
			roomId: 'host-room',
			status: 'connected',
			synced: true,
		});
		expect(last?.people).toEqual([
			expect.objectContaining({ clientId: host.awareness.clientID, name: 'Alice', self: true }),
			expect.objectContaining({ clientId: 99, name: 'Bo', self: false }),
		]);
		const calls = onChange.mock.calls.length;
		act(() => root.render(<Probe config={config} onChange={onChange} />));
		expect(onChange.mock.calls.length).toBe(calls);
		act(() => root.render(<Probe onChange={onChange} />));
		expect(onChange.mock.calls.at(-1)?.[0]).toMatchObject({
			active: false,
			status: 'off',
			people: [],
		});
	});
});
