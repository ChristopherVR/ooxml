/**
 * broadcast-helpers.ts: framework-agnostic helpers for the Broadcast dialog,
 * shared by the React, Vue and Angular bindings.
 *
 * A broadcast is a one-way collaboration session: the presenter drives slide
 * navigation and viewers follow along via a shareable link. These helpers
 * cover the testable, framework-agnostic parts of each binding's broadcast
 * dialog: room-id generation and form validation. The viewer follow-link is
 * `buildBroadcastViewerUrl` in `share-session.ts`, beside the invitation link
 * it shares its format with.
 *
 * No `any`; all regexes use the `/u` flag; no `String.prototype.replaceAll`,
 * no regex named-capture-groups (Angular vendors this source under its
 * ng-packagr lib target). Room ids are generated with `secureRandomToken`
 * (cryptographically strong), read inside the generator at call time, never
 * at module eval.
 */

import { createShareSecret } from 'ooxml-core/collab';

import type { CollaborationTransport } from '../types';
import { secureRandomToken } from './secure-random';

/** Default y-websocket server URL used when no default is supplied. */
export const DEFAULT_BROADCAST_SERVER_URL = 'ws://localhost:1234';

/** Optional seed values for the broadcast start form. */
export interface BroadcastDefaults {
	roomId?: string;
	serverUrl?: string;
}

/** The configuration emitted when a broadcast starts. */
export interface BroadcastConfig {
	roomId: string;
	serverUrl: string;
	/** Derived from serverUrl: empty server means peer-to-peer (webrtc). */
	transport?: CollaborationTransport;
	/** A peer-to-peer broadcast's session secret (see `CollaborationConfig.secret`). */
	secret?: string;
}

/**
 * Resolve the transport implied by a server-URL form field: a blank server
 * URL selects the serverless y-webrtc transport (peers meet via WebRTC
 * signaling plus same-browser BroadcastChannel); anything else is a
 * y-websocket server URL.
 */
export function resolveTransportForServerUrl(serverUrl: string): CollaborationTransport {
	return serverUrl.trim().length === 0 ? 'webrtc' : 'websocket';
}

/** Generate a fresh, broadcast-scoped room id (`broadcast-<suffix>`). */
export function generateBroadcastRoomId(): string {
	return `broadcast-${secureRandomToken(8)}`;
}

/**
 * Seed the start form from the (optional) defaults, generating a fresh room id
 * when none is supplied and falling back to the default server URL.
 */
export function seedBroadcastFields(defaults?: BroadcastDefaults): BroadcastConfig {
	return {
		roomId: defaults?.roomId ?? generateBroadcastRoomId(),
		serverUrl: defaults?.serverUrl ?? DEFAULT_BROADCAST_SERVER_URL,
	};
}

/**
 * Whether the form can start: the room id is required; the server URL may be
 * blank, which selects the serverless webrtc transport.
 */
export function canStartBroadcast(fields: BroadcastConfig): boolean {
	return fields.roomId.trim().length > 0;
}

/**
 * Assemble a {@link BroadcastConfig} from the (trimmed) form fields, or `null`
 * when incomplete. A blank server URL yields `transport: 'webrtc'` with a fresh
 * session secret.
 */
export function buildBroadcastConfig(fields: BroadcastConfig): BroadcastConfig | null {
	if (!canStartBroadcast(fields)) {
		return null;
	}
	const serverUrl = fields.serverUrl.trim();
	const transport = resolveTransportForServerUrl(serverUrl);
	return {
		roomId: fields.roomId.trim(),
		serverUrl,
		transport,
		...(transport === 'webrtc' ? { secret: createShareSecret() } : {}),
	};
}

/** Whether the runtime exposes a usable async clipboard write API. */
export function canUseClipboard(nav: Navigator | undefined): boolean {
	return (
		typeof nav !== 'undefined' &&
		nav.clipboard !== undefined &&
		typeof nav.clipboard.writeText === 'function'
	);
}
