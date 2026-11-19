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
 * the page's other parameters and any fragment that is not a session secret.
 */
export function showSessionInAddressBar(
	config: SessionLink | null,
	param: ShareLinkParam = 'room',
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
