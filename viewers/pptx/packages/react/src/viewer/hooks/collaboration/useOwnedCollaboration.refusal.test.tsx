// @vitest-environment happy-dom
/**
 * A session the React viewer opens itself (a `provider` config or a same-browser room) reports an
 * update it refused as `refusal`, so the viewer can say why sharing stopped, as every other
 * binding's session controller does.
 */
import { createMemoryHub, transportProvider } from 'ooxml-core/collab';
import type { CollaborationConfig } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, expect, test, vi } from 'vitest';

import type { CollaborationContextValue } from './types';
import { useCollaborativeState } from './useCollaborativeState';

const transports = vi.hoisted(() => ({ websocket: vi.fn(), webrtc: vi.fn() }));
vi.mock(import('y-websocket'), () => ({ WebsocketProvider: transports.websocket }));
vi.mock(import('y-webrtc'), () => ({ WebrtcProvider: transports.webrtc }));

/** An update whose struct count no well-formed update of its size can hold. */
const UNREADABLE = new Uint8Array([0, 2, 5, 1, 1, 1, 0, 0x1f]);

let current: CollaborationContextValue | null = null;
function Probe({ config }: { config: CollaborationConfig }) {
	current = useCollaborativeState({ config, canvasWidth: 960, canvasHeight: 540 });
	return null;
}

const settle = () =>
	new Promise((resolve) => {
		setTimeout(resolve, 50);
	});

afterEach(() => {
	current = null;
	vi.unstubAllGlobals();
});

test('reports a refusal for a provider config the viewer opens itself', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const hub = createMemoryHub();
	const root = createRoot(document.createElement('div'));
	const config = {
		roomId: 'room-1',
		serverUrl: '',
		userName: 'Ada',
		provider: transportProvider({ transport: hub.createTransport('room-1') }),
	} as unknown as CollaborationConfig;
	await act(async () => {
		root.render(<Probe config={config} />);
		await settle();
	});
	expect(current?.status).toBe('connected');
	expect(current?.refusal ?? null).toBeNull();

	const peer = hub.createTransport('room-1');
	peer.connect({ open: () => {}, close: () => {}, message: () => {}, error: () => {} });
	await act(async () => {
		peer.send(UNREADABLE);
		await settle();
	});
	expect(current?.status).toBe('error');
	expect(current?.refusal).toBe('malformed');
	act(() => root.unmount());
});
