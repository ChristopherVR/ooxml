/**
 * Framework-neutral Share dialog form, invitation-link, and join parsing.
 * Every UI binding emits the same CollaborationConfig shape from these helpers.
 */
import {
	createShareSecret,
	readShareInvitation,
	shareInvitation,
	type ShareLinkParam,
} from 'ooxml-core/collab';
import type { TranslationKey } from '../i18n';
import type { CollaborationConfig } from '../types';
import { resolveTransportForServerUrl } from './broadcast-helpers';
import { collaborationFailureMessageKey } from './collaboration-failure';
import { isValidRoomId } from './collaboration-presence';

/** A fresh peer-to-peer session secret (`createShareSecret` in `ooxml-core/collab`). */
export { createShareSecret } from 'ooxml-core/collab';

export interface ShareSessionDefaults {
	roomId?: string;
	userName?: string;
	serverUrl?: string;
}

export interface ShareSessionFields {
	roomId: string;
	userName: string;
	serverUrl: string;
}

export interface JoinSessionFields {
	/** A complete invitation URL or a room ID. */
	invitation: string;
	userName: string;
	/** Used when `invitation` is a room ID rather than a complete URL. */
	serverUrl: string;
}

export function seedShareSessionFields(defaults?: ShareSessionDefaults): ShareSessionFields {
	return {
		roomId: defaults?.roomId ?? '',
		userName: defaults?.userName ?? '',
		serverUrl: defaults?.serverUrl ?? '',
	};
}

export function canCreateCollaborationSession(fields: ShareSessionFields): boolean {
	return isValidRoomId(fields.roomId.trim()) && fields.userName.trim().length > 0;
}

export function buildCreateCollaborationConfig(
	fields: ShareSessionFields,
): CollaborationConfig | null {
	if (!canCreateCollaborationSession(fields)) {
		return null;
	}
	const serverUrl = fields.serverUrl.trim();
	const transport = resolveTransportForServerUrl(serverUrl);
	return {
		roomId: fields.roomId.trim(),
		userName: fields.userName.trim(),
		serverUrl,
		transport,
		...(transport === 'webrtc' ? { secret: createShareSecret() } : {}),
		role: 'collaborator',
		sessionIntent: 'create',
	};
}

/**
 * Parse an invitation produced by any framework binding (or any product: the links are the
 * shared `ooxml-core/collab` format). Complete links carry the standard `room`, `server`,
 * `transport`, and optional `signaling` query parameters, and a peer-to-peer link its secret in
 * the `#key=` fragment (without one it names no session). A bare room ID uses the supplied
 * fallback server, and names no session without one.
 */
export function buildJoinCollaborationConfig(
	fields: JoinSessionFields,
): CollaborationConfig | null {
	const userName = fields.userName.trim();
	if (!userName) {
		return null;
	}
	const invitation = readShareInvitation(fields.invitation, fields.serverUrl);
	if (invitation.kind !== 'session') {
		return null;
	}
	const { target } = invitation;
	if (target.transport !== 'webrtc' && !target.serverUrl) {
		// A room code with no server would mean peer-to-peer, which needs the secret only a
		// complete invitation link carries.
		return null;
	}
	return {
		roomId: target.roomId,
		userName,
		serverUrl: target.serverUrl,
		transport: target.transport ?? 'websocket',
		...(target.transport === 'webrtc'
			? {
					...(target.signaling?.length ? { signaling: target.signaling } : {}),
					secret: target.secret,
				}
			: {}),
		role: 'collaborator',
		sessionIntent: 'join',
	};
}

export function canJoinCollaborationSession(fields: JoinSessionFields): boolean {
	return buildJoinCollaborationConfig(fields) !== null;
}

/**
 * The reason the Share dialog's form shows above its fields: on the join tab, why the typed
 * invitation cannot join (a peer-to-peer link without its `#key=` secret); otherwise
 * `sessionFailureKey`, why the last session failed to start.
 */
export function shareFormFailureKey<K extends string>(
	form: { mode: 'create' | 'join'; invitation: string },
	sessionFailureKey: K | null,
): TranslationKey | K | null {
	if (form.mode === 'join' && readShareInvitation(form.invitation).kind === 'missing-secret') {
		return collaborationFailureMessageKey('error', 'missing-secret');
	}
	return sessionFailureKey;
}

/**
 * The invitation link for `config` (under `param`, `room` or a broadcast's `broadcast`). A session
 * without a server is peer-to-peer in PowerPoint, so its link says so and carries the secret in
 * its fragment; without a secret there is no session to join, so only the room code is returned.
 */
export function buildCollaborationShareUrl(
	config: Pick<CollaborationConfig, 'roomId' | 'serverUrl' | 'transport' | 'signaling' | 'secret'>,
	location?: { origin: string; pathname: string },
	param: ShareLinkParam = 'room',
): string {
	const webrtc = config.transport === 'webrtc' || !config.serverUrl.trim();
	if (!webrtc) {
		return shareInvitation({ roomId: config.roomId, serverUrl: config.serverUrl }, location, param);
	}
	if (!config.secret) {
		return config.roomId;
	}
	return shareInvitation(
		{
			roomId: config.roomId,
			serverUrl: '',
			transport: 'webrtc',
			...(config.signaling?.length ? { signaling: config.signaling } : {}),
			secret: config.secret,
		},
		location,
		param,
	);
}

/**
 * The shareable viewer follow-link for a broadcast: its invitation link under the `broadcast`
 * parameter. Returns just the room id when no `origin`/`pathname` are available (non-browser
 * environments) or a peer-to-peer broadcast has no secret.
 */
export function buildBroadcastViewerUrl(
	config: Pick<CollaborationConfig, 'roomId' | 'serverUrl' | 'transport' | 'signaling' | 'secret'>,
	location?: { origin: string; pathname: string },
): string {
	return buildCollaborationShareUrl(config, location, 'broadcast');
}
