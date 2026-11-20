/**
 * The page address as a live invitation. Every pptx demo joins the session its `?room=` or
 * `?broadcast=` link names and, once a session starts, shows that session's invitation in the
 * address bar. Both go through the shared invitation format (`ooxml-core/collab`), so a
 * peer-to-peer link carries its secret in the `#key=` fragment exactly as every binding's Share
 * dialog link does.
 */
import { readShareInvitation, shareInvitation } from 'ooxml-core/collab';
import type { ShareLinkParam, ShareTarget } from 'ooxml-core/collab';

/** The session fields of a binding's `CollaborationConfig` that an invitation carries. */
export interface SessionLink {
	roomId: string;
	serverUrl: string;
	transport?: 'websocket' | 'webrtc';
	signaling?: string[];
	secret?: string;
}

const LINK_PARAMS = ['room', 'broadcast', 'server', 'transport', 'signaling'];

/**
 * The session the page address names under `param`, with `fallbackServerUrl` for a link that
 * names no server, or null. A peer-to-peer link without its `#key=` secret (an older link kept it
 * in the query) still names its room, without a secret: the viewer's session then fails to start
 * and shows its translated missing-key reason instead of the page doing nothing.
 */
export function sessionFromPage(
	param: ShareLinkParam,
	fallbackServerUrl: string,
): SessionLink | null {
	const invitation = readShareInvitation(window.location.href, fallbackServerUrl, param);
	if (invitation.kind === 'none') {
		return null;
	}
	if (invitation.kind === 'missing-secret') {
		return { roomId: invitation.roomId, serverUrl: '', transport: 'webrtc' };
	}
	const { target } = invitation;
	if (target.transport !== 'webrtc') {
		return { roomId: target.roomId, serverUrl: target.serverUrl };
	}
	return {
		roomId: target.roomId,
		serverUrl: '',
		transport: 'webrtc',
		...(target.signaling ? { signaling: target.signaling } : {}),
		secret: target.secret,
	};
}

/** What the page address asks this tab to join, with the role and intent it joins under. */
export interface PageInvitation {
	link: SessionLink;
	/** A `?broadcast=` link joins the audience, which follows the presenter's slides. */
	role?: 'viewer';
	/**
	 * A page that brings its own deck (`?sample=1`) starts the session; any other link joins one,
	 * so the placeholder deck it opens with never seeds the room ahead of the sharer's deck.
	 */
	sessionIntent: 'create' | 'join';
}

/**
 * The session the page address invites this tab to, or null. A `?broadcast=` link wins over a
 * `?room=` link (a crafted address naming both joins the one-way, read-only session).
 */
export function invitationFromPage(fallbackServerUrl: string): PageInvitation | null {
	const sessionIntent =
		new URLSearchParams(window.location.search).get('sample') === '1' ? 'create' : 'join';
	const broadcast = sessionFromPage('broadcast', fallbackServerUrl);
	if (broadcast) {
		return { link: broadcast, role: 'viewer', sessionIntent };
	}
	const room = sessionFromPage('room', fallbackServerUrl);
	return room ? { link: room, sessionIntent } : null;
}

/** A session without a server is peer-to-peer in PowerPoint; it has a link only with its secret. */
function shareTarget(config: SessionLink): ShareTarget | null {
	if (config.transport !== 'webrtc' && config.serverUrl.trim()) {
		return { roomId: config.roomId, serverUrl: config.serverUrl };
	}
	if (!config.secret) {
		return null;
	}
	return {
		roomId: config.roomId,
		serverUrl: '',
		transport: 'webrtc',
		...(config.signaling?.length ? { signaling: config.signaling } : {}),
		secret: config.secret,
	};
}

/**
 * Show `config`'s invitation under `param` in the address bar (or remove it, for `null`), keeping
 * the page's other parameters and any fragment that is not a session secret. A broadcast (the
 * presenter's `owner` session, or the audience's `viewer` one) shows under `?broadcast=`, so the
 * address a presenter copies invites an audience.
 */
export function showSessionInAddressBar(
	config: (SessionLink & { role?: string }) | null,
	param: ShareLinkParam = config?.role === 'owner' || config?.role === 'viewer'
		? 'broadcast'
		: 'room',
): void {
	const url = new URL(window.location.href);
	for (const key of LINK_PARAMS) {
		url.searchParams.delete(key);
	}
	if (new URLSearchParams(url.hash.slice(1)).has('key')) {
		url.hash = '';
	}
	const target = config ? shareTarget(config) : null;
	if (target && url.origin !== 'null') {
		const invitation = new URL(shareInvitation(target, url, param));
		invitation.searchParams.forEach((value, key) => url.searchParams.set(key, value));
		if (invitation.hash) {
			url.hash = invitation.hash;
		}
	}
	window.history.replaceState({}, '', url.toString());
}
