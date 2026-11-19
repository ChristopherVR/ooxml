// The Share dialog's logic every product uses: the fields a user fills in, the invitation link a
// session hands out, and reading one back (a complete link or a bare room code) when joining.
// DOM-free; the dialog itself is `office-ui-share-dialog` in ooxml-ui.
//
// Generalised from the PowerPoint viewer's `share-session.ts` (now in ooxml-ui/pptx, which
// delegates here) so Word, Excel and Visio share and join the same way; see PROVENANCE.md.
import { encodeBase64 } from '../digest/base64';
import { sha256 } from '../digest/sha256';
import type { CollaborationOptions } from './embedding';
import { isValidRoomId } from './validation';

/** What the Share dialog's start form holds. A blank `serverUrl` means "no server". */
export interface ShareFields {
	roomId: string;
	userName: string;
	serverUrl: string;
}

/** What the Share dialog's join form holds. */
export interface JoinFields {
	/** A complete invitation link or a bare room code. */
	invitation: string;
	userName: string;
	/** Used when `invitation` is a room code, which carries no server. */
	serverUrl: string;
}

/**
 * Where a session lives, as an invitation carries it. Without a `transport`, a blank `serverUrl`
 * is the products' default without a server, a BroadcastChannel between windows of this browser.
 */
export interface ServerShareTarget {
	roomId: string;
	serverUrl: string;
	transport?: undefined;
}

/**
 * A peer-to-peer session through WebRTC signaling (only PowerPoint offers it). The signaling
 * servers are third parties, so the session is keyed by `secret` alone: its hash names the
 * signaling room ({@link webrtcRoomName}) and the secret itself encrypts the signaling. The room
 * name is only what the users see.
 */
export interface WebrtcShareTarget {
	roomId: string;
	serverUrl: string;
	transport: 'webrtc';
	signaling?: string[];
	/** A {@link createShareSecret} value. Links carry it in their fragment, never their query. */
	secret: string;
}

export type ShareTarget = ServerShareTarget | WebrtcShareTarget;

/** The query parameter that names the session: `room`, or PowerPoint's broadcast follow link. */
export type ShareLinkParam = 'room' | 'broadcast';

/** The page an invitation link points at; `window.location` fits. */
export interface ShareLocation {
	origin: string;
	pathname: string;
}

/** A fresh room name another window can type: `prefix-` and eight letters or digits. */
export function suggestRoomId(prefix = 'room'): string {
	const bytes = new Uint8Array(8);
	globalThis.crypto.getRandomValues(bytes);
	const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789';
	let suffix = '';
	for (const byte of bytes) suffix += alphabet[byte % alphabet.length];
	return `${prefix}-${suffix}`;
}

// 16 bytes in unpadded base64url: 21 free characters and a last one holding the final 2 bits.
const SECRET_PATTERN = /^[A-Za-z0-9_-]{21}[AQgw]$/u;
const SECRET_FRAGMENT_KEY = 'key';
const WEBRTC_ROOM_DOMAIN = 'ooxml-webrtc-room:';

/** A fresh 128-bit secret for a peer-to-peer session, in unpadded base64url (22 characters). */
export function createShareSecret(): string {
	const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
	return encodeBase64(bytes).replace(/\+/gu, '-').replace(/\//gu, '_').replace(/=+$/u, '');
}

/** Whether `value` has the shape {@link createShareSecret} produces. */
export function isValidShareSecret(value: unknown): value is string {
	return typeof value === 'string' && SECRET_PATTERN.test(value);
}

/**
 * The signaling room of a peer-to-peer session: the hex SHA-256 of its secret under a fixed
 * prefix, so the signaling server learns neither the secret nor the room name users see.
 */
export function webrtcRoomName(secret: string): string {
	const digest = sha256(new TextEncoder().encode(WEBRTC_ROOM_DOMAIN + secret));
	return Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** The start form's fields from optional defaults. */
export function seedShareFields(defaults?: Partial<ShareFields>): ShareFields {
	return {
		roomId: defaults?.roomId ?? '',
		userName: defaults?.userName ?? '',
		serverUrl: defaults?.serverUrl ?? '',
	};
}

/** Whether the start form can start: a valid room name and a display name. */
export function canStartShare(fields: ShareFields): boolean {
	return isValidRoomId(fields.roomId.trim()) && fields.userName.trim().length > 0;
}

/**
 * What an invitation names: a session, a peer-to-peer room that cannot be joined because the link
 * has no valid `#key=` secret (a link in the older format kept its key in the query, or carries
 * none), or nothing at all (blank, or no valid room name).
 */
export type ShareInvitation =
	| { kind: 'session'; target: ShareTarget }
	| { kind: 'missing-secret'; roomId: string }
	| { kind: 'none' };

const NO_INVITATION: ShareInvitation = { kind: 'none' };

/**
 * Read a join form's invitation: a link's `room` (or `param`), `server`, `transport` and
 * `signaling` parameters, or a bare room code with the fallback server. A peer-to-peer link names
 * a session only with a valid secret in its `#key=` fragment; without one it is `missing-secret`.
 */
export function readShareInvitation(
	invitation: string,
	fallbackServerUrl = '',
	param: ShareLinkParam = 'room',
): ShareInvitation {
	const value = invitation.trim();
	if (!value) return NO_INVITATION;
	let url: URL | null = null;
	try {
		url = new URL(value);
	} catch {
		url = null;
	}
	const roomId = (url ? (url.searchParams.get(param) ?? '') : value).trim();
	if (!isValidRoomId(roomId)) return NO_INVITATION;
	if (url?.searchParams.get('transport') === 'webrtc') {
		const secret = new URLSearchParams(url.hash.slice(1)).get(SECRET_FRAGMENT_KEY);
		if (!isValidShareSecret(secret)) return { kind: 'missing-secret', roomId };
		const signaling = (url.searchParams.get('signaling') ?? '')
			.split(',')
			.map((entry) => entry.trim())
			.filter((entry) => entry.length > 0);
		return {
			kind: 'session',
			target: {
				roomId,
				serverUrl: '',
				transport: 'webrtc',
				...(signaling.length ? { signaling } : {}),
				secret,
			},
		};
	}
	return {
		kind: 'session',
		target: { roomId, serverUrl: (url?.searchParams.get('server') ?? fallbackServerUrl).trim() },
	};
}

/** Whether the join form names a session and a display name. */
export function canJoinShare(fields: JoinFields): boolean {
	return joinShareOptions(fields) !== null;
}

/**
 * The invitation for `target`: a link to `location` with `room` (or `param`) and `server` (or
 * `transport` and `signaling`) parameters, a peer-to-peer secret in the `#key=` fragment, which
 * browsers never send to the page's server, or the bare room code where no absolute link can be
 * made (no location, or an opaque origin such as `file:` or a sandboxed frame). A room code joins
 * through the recipient's own server, so it carries no server, signaling or secret.
 */
export function shareInvitation(
	target: ShareTarget,
	location?: ShareLocation,
	param: ShareLinkParam = 'room',
): string {
	if (!location || location.origin === 'null' || location.origin === '') return target.roomId;
	const url = new URL(location.pathname, location.origin);
	url.searchParams.set(param, target.roomId);
	if (target.transport === 'webrtc') {
		url.searchParams.set('transport', 'webrtc');
		if (target.signaling?.length) url.searchParams.set('signaling', target.signaling.join(','));
		url.hash = new URLSearchParams({ [SECRET_FRAGMENT_KEY]: target.secret }).toString();
	} else if (target.serverUrl.trim()) {
		url.searchParams.set('server', target.serverUrl.trim());
	}
	return url.toString();
}

/**
 * The `collaboration` options for a product editor (Word, Excel, Visio): the room, the server
 * when there is one, the user's name and whether this window starts the session or joins it.
 */
export function shareOptions(
	target: ShareTarget,
	userName: string,
	intent: 'create' | 'join',
): CollaborationOptions {
	const name = userName.trim();
	const server = target.serverUrl.trim();
	return {
		roomId: target.roomId,
		...(server ? { serverUrl: server } : {}),
		...(name ? { user: { name } } : {}),
		intent,
	};
}

/** The options the start form describes, or null while it cannot start. */
export function startShareOptions(fields: ShareFields): CollaborationOptions | null {
	if (!canStartShare(fields)) return null;
	return shareOptions(
		{ roomId: fields.roomId.trim(), serverUrl: fields.serverUrl },
		fields.userName,
		'create',
	);
}

/** The options the join form describes, or null while it names no session or no name. */
export function joinShareOptions(fields: JoinFields): CollaborationOptions | null {
	if (!fields.userName.trim()) return null;
	const invitation = readShareInvitation(fields.invitation, fields.serverUrl);
	return invitation.kind === 'session'
		? shareOptions(invitation.target, fields.userName, 'join')
		: null;
}

/**
 * The room a page was opened to join (its `room` query parameter), so a host can open the Share
 * dialog's join form with an invitation already filled in.
 */
export function invitationFromLocation(href: string): string | null {
	try {
		const url = new URL(href);
		return isValidRoomId(url.searchParams.get('room') ?? '') ? url.toString() : null;
	} catch {
		return null;
	}
}
