// The product mapping seam. The collab area knows nothing about slides, paragraphs or cells: a
// product implements {@link DocumentAdapter} (model to and from Y types) and `bindDocument` supplies
// the lifecycle every viewer needs around it: seed an empty room, adopt a non-empty one, skip echoes
// of local writes, and gate writes until the session allows them.
// New code written for the collab area; the rules come from pptx-viewer's
// `collaboration-external-readiness.ts` and `collaboration-load-origin.ts`. See PROVENANCE.md.
import type * as Y from 'yjs';
import { type LoadOrigin, shouldRoomReplaceLoad } from './policy.js';
import type { CollabSession } from './session.js';

/** Transaction origin used for the adapter's own writes, so observers can skip echoes. */
export const LOCAL_ORIGIN: unique symbol = Symbol('ooxml-core:collab:local');

export interface DocumentAdapter<TModel> {
	/** True when nothing has been seeded into the shared document yet. */
	isEmpty: (doc: Y.Doc) => boolean;
	/** Materialise the shared document as the product model. */
	read: (doc: Y.Doc) => TModel;
	/**
	 * Write `model` into the document inside `doc.transact(fn, origin)`. Prefer reconciling only what
	 * changed over replacing the whole structure so concurrent edits merge per field.
	 */
	write: (doc: Y.Doc, model: TModel, origin: typeof LOCAL_ORIGIN) => void;
	/** Observe changes (deep); return an unsubscribe. The handler receives the transaction origin. */
	observe: (doc: Y.Doc, onChange: (origin: unknown) => void) => () => void;
}

export interface BindingOptions<TModel> {
	/** The product's current local model, used to seed an empty room. */
	getLocalModel: () => TModel;
	/** The shared document changed remotely (or the room replaced the bootstrap model): adopt it. */
	onRemoteModel: (model: TModel) => void;
}

export interface DocumentBinding<TModel> {
	/** Publish a local edit. Returns false (and writes nothing) while writes are not allowed. */
	push: (model: TModel) => boolean;
	/**
	 * A load finished. Returns true when the room replaced it (bootstrap into a non-empty room);
	 * otherwise a user load is published to the room.
	 */
	handleLoad: (origin: LoadOrigin | undefined, model: TModel) => boolean;
	dispose: () => void;
}

export function bindDocument<TModel, P extends object>(
	session: CollabSession<P>,
	adapter: DocumentAdapter<TModel>,
	options: BindingOptions<TModel>,
): DocumentBinding<TModel> {
	let disposed = false;
	const adoptRoom = (): void => options.onRemoteModel(adapter.read(session.doc));

	const stopObserving = adapter.observe(session.doc, (origin) => {
		if (disposed || origin === LOCAL_ORIGIN) return;
		adoptRoom();
	});
	const settle = (): void => {
		if (disposed) return;
		if (adapter.isEmpty(session.doc)) {
			if (session.canWrite()) adapter.write(session.doc, options.getLocalModel(), LOCAL_ORIGIN);
		} else adoptRoom();
	};
	const stopReady = session.on('ready', settle);
	if (session.canWrite()) settle();

	return {
		push: (model) => {
			if (disposed || !session.canWrite()) return false;
			adapter.write(session.doc, model, LOCAL_ORIGIN);
			return true;
		},
		handleLoad: (origin, model) => {
			if (disposed) return false;
			if (shouldRoomReplaceLoad(origin, !adapter.isEmpty(session.doc))) {
				adoptRoom();
				return true;
			}
			if (session.canWrite()) adapter.write(session.doc, model, LOCAL_ORIGIN);
			return false;
		},
		dispose: () => {
			disposed = true;
			stopObserving();
			stopReady();
		},
	};
}
