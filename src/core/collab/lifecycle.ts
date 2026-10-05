// Making a departing peer actually leave: document-teardown listeners and a synchronous
// BroadcastChannel departure notice. No DOM types are required (the window and channel are
// structurally typed and injectable). Extracted from pptx-viewer `collaboration-teardown.ts` and
// `collaboration-departure.ts`; channel and message names are now configurable and default to
// neutral `ooxml-core:` names (pptx-viewer's old names are exported as legacy constants).
// See PROVENANCE.md.
import { type AwarenessStatesLike, removeAwarenessStatesLocally } from './awareness.js';

/** `postMessage` payload an embedding page can send into a frame to make it leave its room first. */
export const COLLAB_LEAVE_MESSAGE = 'ooxml-core:collab-leave';
/** BroadcastChannel shared by every session in the browser. */
export const DEPARTURE_CHANNEL = 'ooxml-core:collab-departures';
/** Names used by pptx-viewer before the extraction; pass them to stay wire-compatible. */
export const LEGACY_PPTX_LEAVE_MESSAGE = 'pptx-viewer:collab-leave';
export const LEGACY_PPTX_DEPARTURE_CHANNEL = 'pptx-viewer:collab-departures';

export interface TeardownEventLike {
	readonly persisted?: boolean;
	readonly data?: unknown;
}
export type TeardownListener = (event: TeardownEventLike) => void;
export interface TeardownWindowLike {
	addEventListener: (type: string, listener: TeardownListener) => void;
	removeEventListener: (type: string, listener: TeardownListener) => void;
}

export interface TeardownOptions {
	/** Clear local awareness and destroy the transport. Called at most once per departure. */
	leave: () => void;
	/** Re-establish the session after a bfcache restore; also opts into leaving on persisted pagehide. */
	rejoin?: () => void;
	/** Borrowed sessions wait for `pagehide`, since `beforeunload` may be cancelled. */
	leaveOnBeforeUnload?: boolean | (() => boolean);
	/** Window to listen on; defaults to the ambient `window` when present. */
	target?: TeardownWindowLike;
	/** The leave `postMessage` payload type; defaults to {@link COLLAB_LEAVE_MESSAGE}. */
	leaveMessage?: string;
}

function isLeaveMessage(data: unknown, expected: string): boolean {
	if (typeof data === 'string') return data === expected;
	return (
		typeof data === 'object' && data !== null && (data as { type?: unknown }).type === expected
	);
}

/**
 * Register `pagehide`, `beforeunload`, `pageshow` and `message` listeners that make the local peer
 * leave its room when the document goes away (tab close, navigation, or an embedder detaching the
 * frame, which fires `pagehide` but never `beforeunload`). Returns a disposer; a no-op outside a
 * browser.
 */
export function registerTeardown(options: TeardownOptions): () => void {
	const target = options.target ?? (globalThis as { window?: TeardownWindowLike }).window;
	if (!target) return () => {};
	const { leave, rejoin } = options;
	const leaveMessage = options.leaveMessage ?? COLLAB_LEAVE_MESSAGE;
	let left = false;
	const doLeave = (): void => {
		if (left) return;
		left = true;
		leave();
	};
	const onPageHide: TeardownListener = (event) => {
		// A bfcache'd page with no way back keeps its session so it still works when restored.
		if (event.persisted === true && !rejoin) return;
		doLeave();
	};
	const onBeforeUnload: TeardownListener = () => {
		const allowed =
			typeof options.leaveOnBeforeUnload === 'function'
				? options.leaveOnBeforeUnload()
				: options.leaveOnBeforeUnload !== false;
		if (allowed) doLeave();
	};
	const onPageShow: TeardownListener = (event) => {
		if (event.persisted !== true || !left || !rejoin) return;
		left = false;
		rejoin();
	};
	const onMessage: TeardownListener = (event) => {
		if (isLeaveMessage(event.data, leaveMessage)) doLeave();
	};
	const listeners = {
		pagehide: onPageHide,
		beforeunload: onBeforeUnload,
		pageshow: onPageShow,
		message: onMessage,
	};
	for (const [type, listener] of Object.entries(listeners)) target.addEventListener(type, listener);
	return () => {
		for (const [type, listener] of Object.entries(listeners))
			target.removeEventListener(type, listener);
	};
}

export interface DepartureNotice {
	channel: string;
	roomId: string;
	clientId: number;
}
export interface DepartureChannelLike {
	postMessage: (message: DepartureNotice) => void;
	close: () => void;
	onmessage: ((event: { data: unknown }) => void) | null;
}
export type DepartureChannelFactory = (name: string) => DepartureChannelLike;

export interface DepartureChannel {
	/** Announce that this client is leaving. Synchronous on purpose (usable from `pagehide`). */
	announce: () => void;
	dispose: () => void;
}

function defaultFactory(name: string): DepartureChannelLike | null {
	const scope = globalThis as { BroadcastChannel?: new (name: string) => DepartureChannelLike };
	return scope.BroadcastChannel ? new scope.BroadcastChannel(name) : null;
}

/**
 * Same-browser departure channel: peers announcing their exit are dropped from `awareness` at once
 * instead of lingering as ghosts until the awareness timeout (y-webrtc broadcasts removals one
 * microtask too late for a document being destroyed). Cross-device peers still rely on the
 * transport. A no-op where `BroadcastChannel` is unavailable.
 */
export function createDepartureChannel(
	roomId: string,
	awareness: AwarenessStatesLike,
	factory?: DepartureChannelFactory,
	channelName: string = DEPARTURE_CHANNEL,
): DepartureChannel {
	const channel = factory ? factory(channelName) : defaultFactory(channelName);
	if (!channel) return { announce: () => {}, dispose: () => {} };
	const selfId = awareness.clientID;
	channel.onmessage = (event) => {
		const data = event.data;
		if (typeof data !== 'object' || data === null) return;
		const notice = data as Partial<DepartureNotice>;
		if (notice.channel !== channelName || notice.roomId !== roomId) return;
		if (typeof notice.clientId !== 'number' || notice.clientId === selfId) return;
		removeAwarenessStatesLocally(awareness, [notice.clientId]);
	};
	let closed = false;
	return {
		announce: () => {
			if (closed || typeof selfId !== 'number') return;
			channel.postMessage({ channel: channelName, roomId, clientId: selfId });
		},
		dispose: () => {
			if (closed) return;
			closed = true;
			channel.onmessage = null;
			channel.close();
		},
	};
}
