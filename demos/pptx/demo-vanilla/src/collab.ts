import type { CollaborationConfig } from 'pptx-vanilla-viewer';

import { invitationFromPage } from '../../shared/collaboration-link';
import type { PageInvitation } from '../../shared/collaboration-link';

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

/** The serverless WebRTC collaboration config that joins `invitation` as `userName`. */
export function buildInvitationConfig(
	{ link, role, sessionIntent }: PageInvitation,
	userName: string,
): CollaborationConfig {
	return {
		...link,
		userName,
		userColor: randomUserColor(),
		sessionIntent,
		...(role ? { role } : {}),
	};
}

/**
 * The peer-to-peer session the current URL invites this tab to
 * (`?room=<id>&transport=webrtc#key=<secret>`, or a `?broadcast=` audience link), or null. A link
 * without its secret names the room without one, so the viewer says why it cannot join.
 */
export function readInvitationFromUrl(): PageInvitation | null {
	const invitation = invitationFromPage('');
	return invitation?.link.transport === 'webrtc' ? invitation : null;
}
