import type { CollabYSession, CollaborationConfig, CollaborationTransport } from 'ooxml-ui/pptx';
import { openCollaborationYSession } from 'ooxml-ui/pptx';

import { collabModules } from './collaboration-modules';
import { createCollabProvider } from './collaboration-provider';

/** One session's document, factories and transport (the shared `CollabYSession`). */
export type CollaborationSession = CollabYSession;

/** The shared session opener with this package's loaders and provider factory. */
export function createCollaborationSession(
	config: CollaborationConfig,
	transport: CollaborationTransport,
): Promise<CollaborationSession> {
	return openCollaborationYSession(config, transport, {
		...collabModules,
		createProvider: (kind, settings, doc) => createCollabProvider(kind, settings, doc),
	});
}
