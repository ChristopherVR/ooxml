/**
 * collaboration-session.ts: the session factory seam of the Svelte controller.
 *
 * The document, factories and transport are the shared
 * `openCollaborationYSession` in `ooxml-ui/pptx`, bound to this package's lazy
 * Yjs loaders. The controller accepts a fake {@link CollabSessionFactory} in
 * tests, so they never load the real `yjs` / `y-webrtc` / `y-websocket`.
 */
import type {
	CollabProviderHandle,
	CollabYSession,
	CollaborationConfig,
	CollaborationTransport,
	ExternalCollaborationSession,
	YDocLike,
	YjsFactories,
	YjsModuleLike,
} from 'ooxml-ui/pptx';
import {
	borrowExternalCollaborationAwareness,
	createYjsFactories,
	openCollaborationYSession,
} from 'ooxml-ui/pptx';

import { collabModules } from './collaboration-modules';

/** A fully-wired collaboration session: the doc, its factories, and transport. */
export interface CollabSession {
	/** The live Y.Doc (structural view used by the shared reconcile helpers). */
	ydoc: YDocLike;
	/** Factories the shared reconcile helpers use to build Y.Map/Array/Text. */
	factories: YjsFactories;
	/** The normalised transport provider handle. */
	provider: CollabProviderHandle;
	/** Tear down the provider and destroy the doc. */
	destroy: () => void;
}

/**
 * Build a live session for the given transport + config. Injectable so tests
 * can substitute an in-memory fake (see `collaboration.svelte.ts`).
 */
export type CollabSessionFactory = (
	transport: CollaborationTransport,
	config: CollaborationConfig,
) => Promise<CollabSession>;

/** Default factory: the shared session opener with this package's loaders. */
export const createDefaultSession: CollabSessionFactory = async (transport, config) => {
	const session = await openCollaborationYSession(config, transport, collabModules);
	return {
		ydoc: session.doc,
		factories: session.factories,
		provider: session.provider!,
		destroy: session.dispose,
	};
};

/** Attach factories and a presence lease, never a transport, to a host's doc. */
export async function createExternalSession(
	session: ExternalCollaborationSession,
): Promise<CollabYSession> {
	const Y = (await collabModules.loadYjs()) as YjsModuleLike;
	const borrowed = borrowExternalCollaborationAwareness(session.awareness);
	return {
		doc: session.doc,
		factories: createYjsFactories(Y, session.doc),
		awareness: borrowed.awareness,
		dispose: borrowed.dispose,
	};
}

/** The shared session shape of a factory-built session. */
export function toSharedSession(session: CollabSession): CollabYSession {
	return {
		doc: session.ydoc,
		factories: session.factories,
		awareness: session.provider.awareness,
		provider: session.provider,
		dispose: session.destroy,
	};
}
