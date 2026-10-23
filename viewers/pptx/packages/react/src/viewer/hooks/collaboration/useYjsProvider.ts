/**
 * useYjsProvider: the React adapter of the shared collaboration session
 * controller (`createCollaborationSessionController` in `ooxml-ui/pptx`).
 *
 * Creates a Y.Doc and connects via one of two transports:
 *  - `websocket` (default): WebSocketProvider against `config.serverUrl`.
 *  - `webrtc`: peer-to-peer WebrtcProvider (no document server); usable from
 *    static hosting. Same-browser tabs sync over BroadcastChannel.
 * The controller owns the connection status, the connect timeout, the
 * first-write gate and the teardown; this hook only maps its state.
 *
 * This hook is intentionally thin: it only manages the transport layer.
 * Application-level collaboration logic lives in useCollaborativeState
 * and usePresenceTracking.
 *
 * @module collaboration/useYjsProvider
 */
import {
	createCollaborationSessionController,
	isMixedContentBlocked,
	registerCollaborationTeardown,
} from 'ooxml-ui/pptx';
import { useEffect, useState } from 'react';
import type { Awareness } from 'y-protocols/awareness';
import type { Doc as YDoc } from 'yjs';

import { useController } from '../useController';
import { collabModules } from './collaboration-modules';
import type { CollaborationConfig, ConnectionStatus } from './types';

// Re-export the upstream type aliases for downstream consumers.
export type { YDoc, Awareness };

// Re-export the shared mixed-content guard so existing importers (and the
// colocated test) keep their `./useYjsProvider` import path.
export { isMixedContentBlocked };

// ---------------------------------------------------------------------------
// Hook input / output
// ---------------------------------------------------------------------------

export interface UseYjsProviderInput {
	/**
	 * Collaboration config, or `undefined` when collaboration is inactive. The
	 * hook is always called (React hook rules) but stays fully dormant, never
	 * opening a transport, while config is absent.
	 */
	config?: CollaborationConfig;
}

export interface UseYjsProviderResult {
	/** Current connection status. */
	status: ConnectionStatus;
	/** The Yjs awareness instance (null until connected). */
	awareness: Awareness | null;
	/** The Yjs document (null until initialised). */
	doc: YDoc | null;
	/** Local awareness client ID. */
	clientId: number | null;
	/**
	 * Whether the provider completed its initial document sync. Local doc
	 * writes should be gated on this so a late joiner never seeds its
	 * bootstrap deck into a room whose content has not arrived yet. Websocket
	 * flips it on the provider's sync event; webrtc on peer sync or after the
	 * grace period (a lone fresh-room peer never receives a sync event).
	 */
	synced: boolean;
	/** Manually retry the connection after a timeout or error. */
	retry: () => void;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

/**
 * Lazily loads `yjs` and the transport, creates a Y.Doc and provider, and
 * tracks the connection lifecycle through the shared session controller.
 * The Yjs packages are dynamically imported so they are fully tree-shaken
 * when collaboration is not enabled.
 *
 * This low-level hook owns built-in transports only. Use useCollaborativeState
 * (or the full/headless viewer) to attach a host-owned external session.
 */
export function useYjsProvider({ config: inputConfig }: UseYjsProviderInput): UseYjsProviderResult {
	// Do not open a second transport for an unsupported external-session input.
	const config = inputConfig?.externalSession ? undefined : inputConfig;
	// The teardown listeners are registered per effect run below, so StrictMode's
	// re-run keeps working on the same controller.
	const [controller] = useState(() =>
		createCollaborationSessionController({ modules: collabModules, registerTeardown: false }),
	);
	const state = useController(controller);
	const signalingKey = config?.signaling?.join('\n');

	useEffect(() => {
		if (!config) {
			controller.stop();
			return;
		}
		void controller.start(config);
		// Unmount is not the only way a session ends: a tab close, a navigation,
		// or an embedding page detaching the viewer's iframe destroys the
		// document without ever running React cleanup, leaving a ghost peer in
		// everyone else's presence list. Leave the room from `pagehide` too.
		const disposeTeardown = registerCollaborationTeardown({
			leave: () => controller.stop(),
			rejoin: () => void controller.retry(),
		});
		return () => {
			disposeTeardown();
			controller.stop();
		};
		// Reconnection is keyed on the transport-affecting fields only, so
		// identity-only changes (e.g. userName) never drop the peer link.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [
		controller,
		config?.roomId,
		config?.serverUrl,
		config?.authToken,
		config?.transport,
		signalingKey,
		Boolean(config),
	]);

	const session = state.active ? controller.getSession() : null;
	return {
		status: state.status,
		awareness: (session?.awareness as Awareness | undefined) ?? null,
		doc: (session?.doc as YDoc | undefined) ?? null,
		clientId: state.clientId,
		synced: state.synced,
		retry: () => {
			void controller.retry();
		},
	};
}
