// Presence model: what a collaborator broadcasts over Yjs awareness (identity plus a product
// specific "where am I" payload), how inbound data is sanitised, and how a presence list is derived.
// The wire format is a single nested `presence` awareness field, compatible with pptx-viewer's.
// Extracted from pptx-viewer `collaboration-presence.ts`; generalised so Word/Excel can supply
// their own location payload instead of slide index and cursor pixels. See PROVENANCE.md.
import { type CollaborationRole, sanitizeRole } from './identity.js';
import {
	clampCoordinate,
	DEFAULT_COLOR,
	sanitizeAvatarUrl,
	sanitizeColor,
	sanitizeIndex,
	sanitizeUserName,
} from './validation.js';

/** The awareness field every client publishes its presence under. */
export const PRESENCE_FIELD = 'presence';
/** Presence entries older than this (ms) are dropped as stale. */
export const STALE_PRESENCE_MS = 30_000;
/** Heartbeat interval (ms): re-publish presence so peers do not time us out. */
export const PRESENCE_HEARTBEAT_MS = 10_000;
/** Minimum interval (ms) between outgoing presence broadcasts. */
export const BROADCAST_THROTTLE_MS = 50;

/** Identity fields fixed for the lifetime of a session. */
export interface PresenceIdentity {
	userName: string;
	userColor: string;
	userAvatar?: string | undefined;
	role?: CollaborationRole | undefined;
}

/** A sanitised remote collaborator: identity plus product specific payload `T`. */
export type RemotePresence<T extends object = object> = PresenceIdentity & {
	clientId: number;
	lastUpdated: string;
} & T;

/** Sanitise the identity part of a raw awareness presence entry. */
export function sanitizeIdentity(raw: Record<string, unknown>): PresenceIdentity {
	return {
		userName: sanitizeUserName(raw.userName),
		userColor: sanitizeColor(raw.userColor, DEFAULT_COLOR),
		userAvatar: sanitizeAvatarUrl(raw.userAvatar),
		role: sanitizeRole(raw.role),
	};
}

/** True when `lastUpdated` is parseable and within `staleMs` of `now`. */
export function isPresenceFresh(
	lastUpdated: string,
	now: number = Date.now(),
	staleMs: number = STALE_PRESENCE_MS,
): boolean {
	const updatedAt = new Date(lastUpdated).getTime();
	return !Number.isNaN(updatedAt) && now - updatedAt <= staleMs;
}

export interface DerivePresenceOptions {
	now?: number;
	staleMs?: number;
}

/**
 * Derive the remote presence list from an awareness state map: skips the local client and entries
 * without a `presence` object, sanitises identity, lets `sanitizePayload` validate the product
 * payload (return `null` to drop the entry), and drops stale entries.
 */
export function derivePresence<T extends object>(
	states: ReadonlyMap<number, Record<string, unknown>>,
	localClientId: number | undefined,
	sanitizePayload: (raw: Record<string, unknown>) => T | null,
	options: DerivePresenceOptions = {},
): RemotePresence<T>[] {
	const { now = Date.now(), staleMs = STALE_PRESENCE_MS } = options;
	const users: RemotePresence<T>[] = [];
	for (const [clientId, state] of states) {
		if (clientId === localClientId) continue;
		const raw = state?.[PRESENCE_FIELD];
		if (!raw || typeof raw !== 'object' || Array.isArray(raw)) continue;
		const record = raw as Record<string, unknown>;
		const payload = sanitizePayload(record);
		if (!payload) continue;
		const lastUpdated = typeof record.lastUpdated === 'string' ? record.lastUpdated : '';
		if (!isPresenceFresh(lastUpdated, now, staleMs)) continue;
		users.push({ ...sanitizeIdentity(record), clientId, lastUpdated, ...payload });
	}
	return users;
}

// ---------------------------------------------------------------------------
// Canvas presence: the payload slide/sheet style products broadcast.
// ---------------------------------------------------------------------------

/** Cursor position in unscaled canvas pixels plus the active page (slide) and selection. */
export interface CanvasPayload {
	activeSlideIndex: number;
	cursorX: number;
	cursorY: number;
	selectedElementId?: string | undefined;
}

export type CanvasPresence = RemotePresence<CanvasPayload>;

/** Sanitise the canvas payload, clamping the cursor to the canvas bounds plus a small margin. */
export function sanitizeCanvasPayload(
	raw: Record<string, unknown>,
	canvasWidth: number,
	canvasHeight: number,
): CanvasPayload {
	return {
		activeSlideIndex: sanitizeIndex(raw.activeSlideIndex),
		cursorX: clampCoordinate(raw.cursorX, 0, canvasWidth),
		cursorY: clampCoordinate(raw.cursorY, 0, canvasHeight),
		selectedElementId:
			typeof raw.selectedElementId === 'string' ? raw.selectedElementId.slice(0, 128) : undefined,
	};
}

/** `derivePresence` specialised to the canvas payload. */
export function deriveCanvasPresence(
	states: ReadonlyMap<number, Record<string, unknown>>,
	localClientId: number | undefined,
	canvasWidth: number,
	canvasHeight: number,
	options: DerivePresenceOptions = {},
): CanvasPresence[] {
	return derivePresence(
		states,
		localClientId,
		(raw) => sanitizeCanvasPayload(raw, canvasWidth, canvasHeight),
		options,
	);
}

/** A remote collaborator's cursor, render ready. */
export interface RemoteCursor {
	clientId: number | string;
	userName: string;
	color: string;
	x: number;
	y: number;
	selectionIds?: string[];
}

/** Project canvas presence to cursors, optionally only those on `activeSlideIndex`. */
export function presenceToCursors(
	presence: readonly CanvasPresence[],
	activeSlideIndex?: number,
): RemoteCursor[] {
	const cursors: RemoteCursor[] = [];
	for (const user of presence) {
		if (activeSlideIndex !== undefined && user.activeSlideIndex !== activeSlideIndex) continue;
		const cursor: RemoteCursor = {
			clientId: user.clientId,
			userName: user.userName,
			color: user.userColor,
			x: user.cursorX,
			y: user.cursorY,
		};
		if (user.selectedElementId) cursor.selectionIds = [user.selectedElementId];
		cursors.push(cursor);
	}
	return cursors;
}

/** Coerce an unknown awareness value into a string-id array, dropping non-strings. */
export function asSelectionIds(value: unknown): string[] {
	return Array.isArray(value)
		? value.filter((entry): entry is string => typeof entry === 'string')
		: [];
}
