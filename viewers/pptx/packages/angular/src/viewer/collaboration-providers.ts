/**
 * collaboration-providers.ts: the transport seam of the Angular
 * `CollaborationService`.
 *
 * The Y.Doc, its factories and the y-websocket / y-webrtc provider are the
 * shared `openCollaborationYSession` in `ooxml-ui/pptx`, bound here to this
 * package's lazy loaders. The Yjs packages are dynamically imported with plain
 * string-literal specifiers (matching the other bindings) so they are fully
 * tree-shaken when collaboration is unused, while still resolving under a
 * consuming app's Vite dev server: an `@vite-ignore`d or variable specifier
 * skips Vite's dev-time bare-specifier rewrite, which then fails to resolve
 * natively in the browser.
 */

import type {
	CollabYjsModules,
	CollabYSession,
	CollaborationConfig,
	CollaborationTransport,
} from 'ooxml-ui/pptx';
import { openCollaborationYSession } from 'ooxml-ui/pptx';

/** Minimal awareness surface used by the service. */
export interface AwarenessLike {
	clientID?: number;
	/** Passing `null` withdraws the local presence (see `clearLocalAwareness`). */
	setLocalState?: (state: null) => void;
	setLocalStateField: (field: string, value: unknown) => void;
	getStates: () => Map<number, Record<string, unknown>>;
	on: (event: string, cb: () => void) => void;
	off?: (event: string, cb: () => void) => void;
}

/** This package's loaders for the optional Yjs packages. */
export const collabModules: CollabYjsModules = {
	loadYjs: () => import('yjs'),
	loadWebsocket: () => import('y-websocket'),
	loadWebrtc: () => import('y-webrtc'),
};

/** Open the session's document and transport (or attach to a host's session). */
export function openCollaborationSession(
	config: CollaborationConfig,
	transport: CollaborationTransport,
): Promise<CollabYSession> {
	return openCollaborationYSession(config, transport, collabModules);
}
