// Whole-package document adapter: the shared document holds the latest saved package bytes (any
// OPC format: VSDX, DOCX, PPTX, XLSX). It is the coarsest correct mapping, for products whose
// edits produce a new package rather than a model a CRDT can merge field by field. Concurrent
// changes do not merge: the last package written wins. New code written for the collab area.
import type * as Y from 'yjs';
import type { DocumentAdapter } from './binding';

export interface PackageAdapterOptions {
	/** Y.Map name. Default `package`. */
	name?: string;
	/** Refuse to publish or adopt packages larger than this. Default 32 MiB. */
	maxBytes?: number;
}

export const DEFAULT_MAX_PACKAGE_BYTES = 32 * 1024 * 1024;

/** One shared package: its bytes and a revision counter peers can display or compare. */
export interface SharedPackage {
	bytes: Uint8Array;
	revision: number;
}

export function packageAdapter(
	options: PackageAdapterOptions = {},
): DocumentAdapter<SharedPackage> {
	const name = options.name ?? 'package';
	const maxBytes = options.maxBytes ?? DEFAULT_MAX_PACKAGE_BYTES;
	const map = (doc: Y.Doc) => doc.getMap<unknown>(name);
	const check = (bytes: unknown): Uint8Array => {
		if (!(bytes instanceof Uint8Array)) throw new Error('The shared package is not binary data.');
		if (bytes.byteLength > maxBytes)
			throw new Error(`The shared package exceeds ${Math.round(maxBytes / 1024 / 1024)} MiB.`);
		return bytes;
	};
	return {
		isEmpty: (doc) => !(map(doc).get('bytes') instanceof Uint8Array),
		read: (doc) => {
			const shared = map(doc);
			const revision = shared.get('revision');
			return {
				bytes: check(shared.get('bytes')).slice(),
				revision: typeof revision === 'number' && Number.isSafeInteger(revision) ? revision : 0,
			};
		},
		write: (doc, model, origin) => {
			const bytes = check(model.bytes);
			doc.transact(() => {
				const shared = map(doc);
				const current = shared.get('revision');
				const revision = Math.max(model.revision, typeof current === 'number' ? current : 0) + 1;
				shared.set('bytes', bytes.slice());
				shared.set('revision', revision);
			}, origin);
		},
		observe: (doc, onChange) => {
			const shared = map(doc);
			const handler = (event: Y.YMapEvent<unknown>) => {
				// Bytes and revision change in one transaction; report it once.
				if (event.keysChanged.has('bytes')) onChange(event.transaction.origin);
			};
			shared.observe(handler);
			return () => shared.unobserve(handler);
		},
	};
}
