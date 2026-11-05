/**
 * Framework-neutral Share dialog form, invitation-link, and join parsing.
 * Every UI binding emits the same CollaborationConfig shape from these helpers.
 */
import { parseShareInvitation, shareInvitation } from 'ooxml-core/collab';
import type { CollaborationConfig } from '../types';
import { resolveTransportForServerUrl } from './broadcast-helpers';
import { isValidRoomId } from './collaboration-presence';

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
	return {
		roomId: fields.roomId.trim(),
		userName: fields.userName.trim(),
		serverUrl,
		transport: resolveTransportForServerUrl(serverUrl),
		role: 'collaborator',
		sessionIntent: 'create',
	};
}

/**
 * Parse an invitation produced by any framework binding (or any product: the links are the
 * shared `ooxml-core/collab` format). Complete links carry the standard `room`, `server`,
 * `transport`, and optional `signaling` query parameters. A bare room ID uses the supplied
 * fallback server.
 */
export function buildJoinCollaborationConfig(
	fields: JoinSessionFields,
): CollaborationConfig | null {
	const userName = fields.userName.trim();
	if (!userName) {
		return null;
	}
	const target = parseShareInvitation(fields.invitation, fields.serverUrl);
	if (!target) {
		return null;
	}
	const transport =
		target.transport === 'webrtc' ? 'webrtc' : resolveTransportForServerUrl(target.serverUrl);
	return {
		roomId: target.roomId,
		userName,
		serverUrl: transport === 'webrtc' ? '' : target.serverUrl,
		transport,
		...(target.signaling?.length ? { signaling: target.signaling } : {}),
		role: 'collaborator',
		sessionIntent: 'join',
	};
}

export function canJoinCollaborationSession(fields: JoinSessionFields): boolean {
	return buildJoinCollaborationConfig(fields) !== null;
}

export function buildCollaborationShareUrl(
	config: Pick<CollaborationConfig, 'roomId' | 'serverUrl' | 'transport' | 'signaling'>,
	location?: { origin: string; pathname: string },
): string {
	// A session without a server is peer-to-peer in PowerPoint, so its link says so.
	const webrtc = config.transport === 'webrtc' || !config.serverUrl.trim();
	return shareInvitation(
		webrtc
			? {
					roomId: config.roomId,
					serverUrl: '',
					transport: 'webrtc',
					...(config.signaling?.length ? { signaling: config.signaling } : {}),
				}
			: { roomId: config.roomId, serverUrl: config.serverUrl },
		location,
	);
}
