// Collaborator identity: roles, colours, labels, initials and collision-resistant id generation.
// Extracted from pptx-viewer `collaboration-presence.ts` / `collaboration-active-session.ts` and
// docx-viewer `collaboration-identity.ts` / `presence.ts` (palette); see PROVENANCE.md.
import { DEFAULT_COLOR, isValidId } from './validation.js';

/** Collaboration role within a session. */
export type CollaborationRole = 'owner' | 'collaborator' | 'viewer';

export const COLLABORATION_ROLES: readonly CollaborationRole[] = [
	'owner',
	'collaborator',
	'viewer',
];

export function sanitizeRole(value: unknown): CollaborationRole | undefined {
	return COLLABORATION_ROLES.includes(value as CollaborationRole)
		? (value as CollaborationRole)
		: undefined;
}

/** Whether a role may write to the shared document. */
export function canWrite(role: CollaborationRole | undefined): boolean {
	return role !== 'viewer';
}

/** Default palette: distinct, legible hues with white-text contrast. */
export const CURSOR_PALETTE: readonly string[] = [
	'#ef4444',
	'#f97316',
	'#eab308',
	'#22c55e',
	'#06b6d4',
	'#3b82f6',
	'#8b5cf6',
	'#ec4899',
];

/** Maximum characters shown in a cursor label before truncation. */
export const MAX_LABEL_CHARS = 20;

/** Deterministically pick a palette colour: the same seed always yields the same colour. */
export function assignUserColor(
	seed: number | string,
	palette: readonly string[] = CURSOR_PALETTE,
): string {
	if (palette.length === 0) return DEFAULT_COLOR;
	let hash = 0;
	const text = String(seed);
	for (let i = 0; i < text.length; i++) hash = (hash * 31 + text.charCodeAt(i)) | 0;
	return palette[Math.abs(hash) % palette.length] ?? DEFAULT_COLOR;
}

/** Truncate to exactly `maxChars` characters including a trailing `...` when too long. */
export function formatCursorLabel(userName: string, maxChars: number = MAX_LABEL_CHARS): string {
	if (userName.length <= maxChars) return userName;
	return `${userName.slice(0, Math.max(0, maxChars - 3))}...`;
}

/** Up to two uppercase initials: first+last for multi-word names, otherwise the first two letters. */
export function getUserInitials(name: string): string {
	const trimmed = name.trim();
	const parts = trimmed.split(/\s+/u).filter(Boolean);
	const first = parts[0];
	const last = parts[parts.length - 1];
	if (parts.length >= 2 && first && last) return (first[0]! + last[0]!).toUpperCase();
	return trimmed.slice(0, 2).toUpperCase();
}

/** A random-enough client id for a session; `crypto.randomUUID` when present. */
export function createClientId(): string {
	const cryptoApi = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
	if (cryptoApi?.randomUUID) return cryptoApi.randomUUID();
	return `c${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
}

function clientKey(clientId: string): string {
	if (!isValidId(clientId)) throw new Error('Invalid collaboration client ID');
	return Array.from(clientId, (character) => character.codePointAt(0)!.toString(16)).join('_');
}

/**
 * Generator of collision-resistant document ids that are unique per client session, so two peers
 * creating a paragraph/shape/row at the same time never mint the same id. `prefix` namespaces the
 * product (`dve` for Word, `pve` for PowerPoint).
 */
export function createIdGenerator(clientId: string, prefix = 'cid'): (kind: string) => string {
	const key = clientKey(clientId);
	let counter = 0;
	return (kind) => {
		const safeKind = kind.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 24) || 'item';
		return `${prefix}-${safeKind}-${key}-${++counter}`;
	};
}

/** The fields a connected-users list needs from a collaborator. */
export interface RosterUserInput {
	clientId: number | string;
	userName: string;
	userColor: string;
	userAvatar?: string | undefined;
}

export interface RosterUser {
	/** `'local'` for the current user, otherwise `String(clientId)`. */
	id: string;
	name: string;
	initials: string;
	color: string;
	avatarUrl?: string;
	isLocal: boolean;
}

/** Local user first, then remote users in their existing order. */
export function buildRoster(params: {
	localUserName: string;
	localUserInitials?: string;
	localUserColor?: string;
	remoteUsers: readonly RosterUserInput[];
}): RosterUser[] {
	const customInitials = params.localUserInitials?.trim();
	const local: RosterUser = {
		id: 'local',
		name: params.localUserName,
		initials: customInitials
			? customInitials.slice(0, 2).toUpperCase()
			: getUserInitials(params.localUserName),
		color: params.localUserColor ?? '#6366f1',
		isLocal: true,
	};
	const remote = params.remoteUsers.map((user): RosterUser => {
		const entry: RosterUser = {
			id: String(user.clientId),
			name: user.userName,
			initials: getUserInitials(user.userName),
			color: user.userColor,
			isLocal: false,
		};
		if (user.userAvatar) entry.avatarUrl = user.userAvatar;
		return entry;
	});
	return [local, ...remote];
}
