import type { DocumentSaveMeta, DocumentStorage } from './storage';

/** Minimal metadata shared by editors. Provider identity and revision fields remain host-owned. */
export interface VersionedFile {
	name: string;
	kind: string;
}
export interface StorageRequestOptions {
	signal?: AbortSignal;
}
export interface StoredDocument<F extends VersionedFile = VersionedFile> {
	file: F;
	bytes: Uint8Array;
}
/** A save MUST compare the supplied revision atomically and reject stale revisions. */
export interface VersionedStorageProvider<F extends VersionedFile> {
	read(file: F, options?: StorageRequestOptions): Promise<StoredDocument<F>>;
	save(file: F, bytes: Uint8Array, options?: StorageRequestOptions): Promise<F>;
	isConflict(error: unknown): boolean;
	saveCopy?(
		file: F,
		bytes: Uint8Array,
		name?: string,
		options?: StorageRequestOptions,
	): Promise<StoredDocument<F>>;
}
export interface StorageConflict<F extends VersionedFile = VersionedFile> {
	base: StoredDocument<F>;
	local: Uint8Array;
	remote: StoredDocument<F>;
	issues?: readonly string[];
}
export interface StorageMergeInput<
	F extends VersionedFile = VersionedFile,
> extends StorageConflict<F> {
	kind: F['kind'];
	signal?: AbortSignal;
}
export type StorageMergeResult =
	| { status: 'merged'; bytes: Uint8Array }
	| { status: 'unresolved'; issues: readonly string[] };
export type StorageMerger<F extends VersionedFile = VersionedFile> = (
	input: StorageMergeInput<F>,
) => StorageMergeResult | Promise<StorageMergeResult>;
export class StorageConflictError<F extends VersionedFile = VersionedFile> extends Error {
	constructor(public readonly conflict: StorageConflict<F>) {
		super(
			'The changes could not be combined automatically. Your edits are retained for resolution.',
		);
		this.name = 'StorageConflictError';
	}
}
export interface StorageSessionOptions<F extends VersionedFile> {
	merge?: StorageMerger<F>;
	/** Default 3 with a merger, 1 without. Integer from 1 to 10. */
	maxMergeAttempts?: number;
	/** Override for host-specific error types. The snapshot remains available on the session. */
	conflictError?: (conflict: StorageConflict<F>) => Error;
	/** Reject a format conversion or other invalid editor save before any provider write. */
	validateSave?: (meta: DocumentSaveMeta, file: Readonly<F>) => void;
	/** Override when file metadata contains nested mutable values. Defaults to structuredClone. */
	cloneFile?: (file: F) => F;
}
export interface StorageSessionSnapshot<F extends VersionedFile> {
	file: F;
	baseBytes: Uint8Array;
	conflict: StorageConflict<F> | null;
}
export interface StorageSession<F extends VersionedFile> {
	getSnapshot(): StorageSessionSnapshot<F>;
	save(bytes: Uint8Array, options?: StorageRequestOptions): Promise<F>;
	reload(options?: StorageRequestOptions): Promise<StoredDocument<F>>;
	/** Rejects if the provider does not implement saveCopy. */
	saveCopy(
		bytes?: Uint8Array,
		name?: string,
		options?: StorageRequestOptions,
	): Promise<StoredDocument<F>>;
	storage: DocumentStorage;
}
