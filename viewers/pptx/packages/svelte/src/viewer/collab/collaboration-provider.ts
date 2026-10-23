/**
 * collaboration-provider.ts: the shared `createCollabProvider` in
 * `ooxml-ui/pptx`, bound to this package's lazy Yjs loaders.
 */
import type {
	AwarenessLike,
	CollabProviderHandle,
	CollaborationConfig,
	CollaborationTransport,
} from 'ooxml-ui/pptx';
import { createCollabProvider as createSharedCollabProvider } from 'ooxml-ui/pptx';

import { collabModules } from './collaboration-modules';

export type { AwarenessLike, CollabProviderHandle };

/** The shared provider factory with this package's loaders. */
export function createCollabProvider(
	transport: CollaborationTransport,
	config: CollaborationConfig,
	doc: unknown,
): Promise<CollabProviderHandle> {
	return createSharedCollabProvider(transport, config, doc, collabModules);
}
