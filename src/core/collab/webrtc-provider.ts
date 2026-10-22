// WebRTC for the embedding contract without a y-webrtc dependency: the host passes the
// `WebrtcProvider` class it installed (as `createWebSocketTransport` takes a `WebSocket`), and this
// wraps each instance through `adaptYjsProvider`. Peer-to-peer rooms never confirm a sync, so the
// session's grace period opens the write gate for the first peer.
import type { Awareness } from 'y-protocols/awareness';
import type * as Y from 'yjs';
import { adaptYjsProvider, type YjsProviderLike } from './external-provider';
import type { ProviderFactory } from './provider';

/** The y-webrtc `WebrtcProvider` constructor, structurally. */
export type WebrtcProviderConstructor = new (
	roomId: string,
	doc: Y.Doc,
	options: { signaling?: string[]; password?: string; awareness?: Awareness },
) => YjsProviderLike;

export interface WebrtcOptions {
	/** `WebrtcProvider` from the host's own `y-webrtc` install. */
	WebrtcProvider: WebrtcProviderConstructor;
	/** Signaling server URLs; y-webrtc's defaults when omitted. */
	signaling?: string[];
	/** Room password: peers without it cannot read the room. */
	password?: string;
}

/** A provider factory that joins the session's room over y-webrtc, sharing its awareness. */
export function webrtcProvider(options: WebrtcOptions): ProviderFactory {
	return ({ doc, awareness, roomId }) =>
		adaptYjsProvider(
			new options.WebrtcProvider(roomId, doc, {
				awareness,
				...(options.signaling ? { signaling: options.signaling } : {}),
				...(options.password ? { password: options.password } : {}),
			}),
		);
}
