/**
 * collaboration-yjs-session.ts: the live Yjs document, its factories and its
 * transport (or a borrowed host session) for one PowerPoint collaboration
 * session, written once instead of in every binding.
 *
 * `yjs` stays an optional peer of the bindings, so the binding passes its own
 * loader (`loadYjs: () => import('yjs')`): the document and the provider must
 * come from the same Yjs instance.
 *
 * @module render/collaboration-yjs-session
 */
import type { CollaborationConfig, CollaborationTransport } from '../types';
import { borrowExternalCollaborationAwareness } from './collaboration-external-session';
import type { AwarenessLike } from './collaboration-presence-publisher';
import type { CollabProviderHandle, CollabTransportModules } from './collaboration-provider';
import { createCollabProvider } from './collaboration-provider';
import type { YDocLike, YjsFactories } from './collaboration-sync';
import { createSnapshotTextPositions } from './collaboration-text-snapshot-positions';

/** The Y.Doc surface the factories and the session need. */
interface YDocInstance {
	on(event: 'beforeObserverCalls', listener: () => void): void;
	off(event: 'beforeObserverCalls', listener: () => void): void;
	destroy(): void;
}

/** `import('yjs')`, structurally. */
export interface YjsModuleLike {
	Doc: new () => unknown;
	Map: new () => unknown;
	Array: new () => unknown;
	Text: new () => unknown;
	snapshot(doc: never): unknown;
	equalSnapshots(a: never, b: never): boolean;
}

/** The binding's lazy loaders: `yjs` plus the two transports (or its own provider factory). */
export interface CollabYjsModules extends CollabTransportModules {
	loadYjs(): Promise<unknown>;
	/** Build the provider; defaults to `createCollabProvider` with these loaders. */
	createProvider?(
		transport: CollaborationTransport,
		config: CollaborationConfig,
		doc: unknown,
	): Promise<CollabProviderHandle>;
}

/** One session's document, factories and transport. */
export interface CollabYSession {
	doc: YDocLike;
	factories: YjsFactories;
	awareness: AwarenessLike;
	/** The owned transport; absent for a borrowed host session. */
	provider?: CollabProviderHandle;
	/** Destroy what this session owns (never the host's document). */
	dispose(): void;
}

/** The factories the shared reconcile helpers use, bound to one document. */
export function createYjsFactories(Y: YjsModuleLike, doc: unknown): YjsFactories {
	const live = doc as YDocInstance;
	return {
		createMap: () => new Y.Map() as ReturnType<YjsFactories['createMap']>,
		createArray: () => new Y.Array() as ReturnType<YjsFactories['createArray']>,
		createText: () => new Y.Text() as ReturnType<YjsFactories['createText']>,
		createTextPositions: (text) =>
			createSnapshotTextPositions(text, {
				read: () => Y.snapshot(doc as never),
				equal: (a, b) => Y.equalSnapshots(a as never, b as never),
				subscribeBeforeObservers: (listener) => {
					live.on('beforeObserverCalls', listener);
					return () => live.off('beforeObserverCalls', listener);
				},
			}),
	};
}

/**
 * Open the session for `config`: a borrowed host session (`externalSession`,
 * no transport, a presence lease) or a fresh Y.Doc with the requested provider.
 */
export async function openCollaborationYSession(
	config: CollaborationConfig,
	transport: CollaborationTransport,
	modules: CollabYjsModules,
): Promise<CollabYSession> {
	const Y = (await modules.loadYjs()) as YjsModuleLike;
	const external = config.externalSession;
	if (external) {
		const borrowed = borrowExternalCollaborationAwareness(external.awareness);
		return {
			doc: external.doc,
			factories: createYjsFactories(Y, external.doc),
			awareness: borrowed.awareness,
			dispose: borrowed.dispose,
		};
	}
	const doc = new Y.Doc() as YDocInstance;
	try {
		const provider = modules.createProvider
			? await modules.createProvider(transport, config, doc)
			: await createCollabProvider(transport, config, doc, modules);
		return {
			doc: doc as unknown as YDocLike,
			factories: createYjsFactories(Y, doc),
			awareness: provider.awareness,
			provider,
			dispose() {
				provider.destroy();
				doc.destroy();
			},
		};
	} catch (error) {
		doc.destroy();
		throw error;
	}
}
