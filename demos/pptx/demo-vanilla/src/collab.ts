import type { CollaborationConfig } from 'pptx-vanilla-viewer';

import { sessionFromPage } from '../../shared/collaboration-link';
import type { SessionLink } from '../../shared/collaboration-link';

/**
 * Collaboration helpers for the vanilla demo, mirroring the Vue demo's
 * `collab.ts` at smaller scale. The vanilla demo ships without a relay server,
 * so it joins over serverless WebRTC (`transport: 'webrtc'`): same-browser tabs
 * meet through BroadcastChannel with no server, which is exactly what a local
 * two-tab test needs.
 */

/** Cryptographically strong base-36 token of `length` characters. */
function secureRandomToken(length: number): string {
	const bytes = new Uint8Array(length);
	crypto.getRandomValues(bytes);
	let out = '';
	for (const b of bytes) {
		out += (b % 36).toString(36);
	}
	return out;
}

/** Generate a stable per-session room id, persisted in sessionStorage. */
export function resolveAutoRoomId(): string {
	const stored = sessionStorage.getItem('pptx-vanilla-room-id');
	if (stored) {
		return stored;
	}
	const id = `session-${secureRandomToken(8)}`;
	sessionStorage.setItem('pptx-vanilla-room-id', id);
	return id;
}

/** Friendly display name from the platform plus a random suffix. */
export function resolveAutoName(): string {
	const ua = navigator.userAgent;
	let platform = 'User';
	if (ua.includes('Win')) {
		platform = 'Windows';
	} else if (ua.includes('Mac')) {
		platform = 'Mac';
	} else if (ua.includes('Linux')) {
		platform = 'Linux';
	}
	return `${platform}-${secureRandomToken(4)}`;
}

/** Random hex colour for a collaboration cursor/label. */
export function randomUserColor(): string {
	const bytes = new Uint8Array(3);
	crypto.getRandomValues(bytes);
	return `#${Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('')}`;
}

/** A serverless WebRTC collaboration config for the session `link` names. */
export function buildRoomConfig(link: SessionLink, userName: string): CollaborationConfig {
	return { ...link, userName, userColor: randomUserColor() };
}

/**
 * The peer-to-peer session the current URL names (`?room=<id>&transport=webrtc#key=<secret>`),
 * or null. A link without its secret names the room without one, so the viewer says why it
 * cannot join.
 */
export function readRoomFromUrl(): SessionLink | null {
	const link = sessionFromPage('room', '');
	return link?.transport === 'webrtc' ? link : null;
}
