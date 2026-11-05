// The Share dialog's logic every product uses: the fields a user fills in, the invitation link a
// session hands out, and reading one back (a complete link or a bare room code) when joining.
// DOM-free; the dialog itself is `office-ui-share-dialog` in ooxml-ui.
//
// Generalised from the PowerPoint viewer's `share-session.ts` (now in ooxml-ui/pptx, which
// delegates here) so Word, Excel and Visio share and join the same way; see PROVENANCE.md.
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
 * Where a session lives, as an invitation carries it. `webrtc` (only PowerPoint offers it) is
 * peer-to-peer through WebRTC signaling; otherwise a blank `serverUrl` is the products' default
 * without a server, a BroadcastChannel between windows of this browser.
 */
export interface ShareTarget {
	roomId: string;
	serverUrl: string;
	transport?: 'webrtc';
	signaling?: string[];
}

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
 * The target a join form names, or null when it names none: a link's `room`, `server`,
 * `transport` and `signaling` parameters, or a bare room code with the fallback server.
 */
export function parseShareInvitation(
	invitation: string,
	fallbackServerUrl = '',
): ShareTarget | null {
	const value = invitation.trim();
	if (!value) return null;
	let url: URL | null = null;
	try {
		url = new URL(value);
	} catch {
		url = null;
	}
	const roomId = (url ? (url.searchParams.get('room') ?? '') : value).trim();
	if (!isValidRoomId(roomId)) return null;
	if (url?.searchParams.get('transport') === 'webrtc') {
		const signaling = (url.searchParams.get('signaling') ?? '')
			.split(',')
			.map((entry) => entry.trim())
			.filter((entry) => entry.length > 0);
		return {
			roomId,
			serverUrl: '',
			transport: 'webrtc',
			...(signaling.length ? { signaling } : {}),
		};
	}
	return { roomId, serverUrl: (url?.searchParams.get('server') ?? fallbackServerUrl).trim() };
}

/** Whether the join form names a session and a display name. */
export function canJoinShare(fields: JoinFields): boolean {
	return (
		fields.userName.trim().length > 0 &&
		parseShareInvitation(fields.invitation, fields.serverUrl) !== null
	);
}

/**
 * The invitation for `target`: a link to `location` with `room` and `server` (or `transport` and
 * `signaling`) parameters, or the bare room code where no absolute link can be made (no location,
 * or an opaque origin such as `file:` or a sandboxed frame). A room code joins through the
 * recipient's own server, so it carries no server or signaling.
 */
export function shareInvitation(target: ShareTarget, location?: ShareLocation): string {
	if (!location || location.origin === 'null' || location.origin === '') return target.roomId;
	const url = new URL(location.pathname, location.origin);
	url.searchParams.set('room', target.roomId);
	const server = target.serverUrl.trim();
	if (target.transport === 'webrtc') {
		url.searchParams.set('transport', 'webrtc');
		if (target.signaling?.length) url.searchParams.set('signaling', target.signaling.join(','));
	} else if (server) {
		url.searchParams.set('server', server);
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
	const target = parseShareInvitation(fields.invitation, fields.serverUrl);
	return target ? shareOptions(target, fields.userName, 'join') : null;
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
