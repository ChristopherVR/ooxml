// The public shape of `<xlsx-editor>` collaboration: what a host passes to `collaboration` /
// `startCollaboration()` and what `collaboration-change` reports. Types only, so importing them
// never loads the Yjs runtime.
import type {
	CollabSession,
	CollaborationRole,
	ConnectionStatus,
	ProviderFactory,
} from 'ooxml-core/collab';
import type { XlsxPresence } from 'ooxml-core/xlsx/collab';

/**
 * How to join a shared workbook. One transport is used, in this order: `session` (host-owned,
 * never destroyed by the editor), `provider`, `serverUrl` (a y-websocket compatible server), else a
 * BroadcastChannel between tabs and windows of this browser.
 */
export interface XlsxCollaborationOptions {
	/** Room name: 1-128 letters, digits, hyphens or underscores. Not needed with `session`. */
	roomId?: string;
	/** Base URL of a y-websocket compatible server (`wss://host/collab`); the room is appended. */
	serverUrl?: string;
	/** A provider factory of your own (WebRTC, an existing provider through `externalProvider`). */
	provider?: ProviderFactory;
	/** An existing session; create it with `sanitizePayload: sanitizeXlsxPresence`. */
	session?: CollabSession<XlsxPresence>;
	/** Who this window is. The name defaults to `authorName`, the colour to a palette colour. */
	user?: { name?: string; color?: string; role?: CollaborationRole };
}

/** One person in the room, as the title bar and `collaboration-change` list them. */
export interface XlsxCollaborator {
	clientId: number;
	name: string;
	color: string;
	self: boolean;
}

/** Collaboration state, reported by `collaborationState` and the `collaboration-change` event. */
export interface XlsxCollaborationState {
	active: boolean;
	roomId: string | undefined;
	/** `off` when not sharing; otherwise the provider's connection status. */
	status: ConnectionStatus | 'off';
	/** The first exchange with the room finished. */
	synced: boolean;
	people: XlsxCollaborator[];
}

export const OFF_STATE: XlsxCollaborationState = Object.freeze({
	active: false,
	roomId: undefined,
	status: 'off',
	synced: false,
	people: [],
});
