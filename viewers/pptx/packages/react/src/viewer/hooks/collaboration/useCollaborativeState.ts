import type { CollaborationWriteBackState } from 'ooxml-core/collab';
import { sanitizeColor } from 'ooxml-ui/pptx';
import { useMemo, useState } from 'react';

/**
 * useCollaborativeState: Composes the Yjs provider and presence tracking
 * into a single hook for the collaboration system.
 *
 * This is the primary hook consumed by the `CollaborationProvider` context.
 * It orchestrates:
 * - Yjs WebSocket connection lifecycle
 * - Presence tracking (broadcast + receive)
 * - Connection status
 *
 * @module collaboration/useCollaborativeState
 */
import type { CollaborationConfig, CollaborationContextValue } from './types';
import { useCollaborationSession } from './useCollaborationSession';
import { useOwnedCollaboration } from './useOwnedCollaboration';
import { usePresenceTracking } from './usePresenceTracking';

// ---------------------------------------------------------------------------
// Input
// ---------------------------------------------------------------------------

export interface UseCollaborativeStateInput {
	/**
	 * Collaboration config, or `undefined` when collaboration is inactive. The
	 * hook is always called so the surrounding provider keeps a stable React
	 * tree shape; it produces a dormant value (no transport, empty presence)
	 * while config is absent.
	 */
	config?: CollaborationConfig;
	canvasWidth: number;
	canvasHeight: number;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useCollaborativeState({
	config: requested,
	canvasWidth,
	canvasHeight,
}: UseCollaborativeStateInput): CollaborationContextValue | null {
	// A `provider` or same-browser room is opened here as a session the viewer owns.
	const ownedSession = useOwnedCollaboration(requested);
	const config = ownedSession.config;
	const userColor = sanitizeColor(config?.userColor, '#6366f1');

	const session = useCollaborationSession({ config: ownedSession.holding ? undefined : config });
	const { awareness, doc, clientId, synced, retry } = session;
	const status = ownedSession.status ?? session.status;

	const { remoteUsers, broadcastPresence } = usePresenceTracking({
		awareness,
		localClientId: clientId,
		userName: config?.userName ?? '',
		userColor,
		userAvatar: config?.userAvatar,
		role: config?.role,
		canvasWidth,
		canvasHeight,
	});

	// The owner's write-back status, reported by the document sync.
	const [writeBack, reportWriteBack] = useState<CollaborationWriteBackState | undefined>();

	// Total connected = remote users + local (if connected)
	const connectedCount = status === 'connected' ? remoteUsers.length + 1 : remoteUsers.length;

	// Collaboration inactive: all the hooks above ran (and stayed dormant) so the
	// provider keeps a stable tree, but expose a null context value so
	// `useCollaboration()` reports "not collaborating", exactly as before.
	return useMemo(
		() =>
			config
				? {
						status,
						remoteUsers,
						broadcastPresence,
						connectedCount,
						config,
						doc,
						synced,
						retry,
						clientId,
						writeBack,
						reportWriteBack,
					}
				: null,
		[
			status,
			remoteUsers,
			broadcastPresence,
			connectedCount,
			config,
			doc,
			synced,
			retry,
			clientId,
			writeBack,
		],
	);
}
