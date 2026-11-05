import { runStorageMerge } from './storage-merge';
import { StorageConflictError } from './versioned-storage';
import type {
	VersionedFile,
	StoredDocument,
	StorageConflict,
	StorageRequestOptions,
	VersionedStorageProvider,
	StorageSession,
	StorageSessionOptions,
} from './versioned-storage';
/**
 * One serialized cloud document identity and its last acknowledged baseline. Bytes from failed
 * saves remain available for conflict recovery. This is storage coordination, not Office merging.
 */
export function createStorageSession<F extends VersionedFile>(
	client: VersionedStorageProvider<F>,
	opened: StoredDocument<F>,
	settings: StorageSessionOptions<F> = {},
): StorageSession<F> {
	const cloneFile: (file: F) => F = settings.cloneFile ?? ((file) => structuredClone(file));
	const clone = (value: StoredDocument<F>): StoredDocument<F> => ({
		file: cloneFile(value.file),
		bytes: value.bytes.slice(),
	});
	const equal = (a: Uint8Array, b: Uint8Array) =>
		a.length === b.length && a.every((value, index) => value === b[index]);
	const cloneConflict = (value: StorageConflict<F>): StorageConflict<F> => ({
		base: clone(value.base),
		local: value.local.slice(),
		remote: clone(value.remote),
		...(value.issues ? { issues: [...value.issues] } : {}),
	});

	const merge = settings.merge;
	const maxRebases = settings.maxMergeAttempts ?? (merge ? 3 : 1);
	if (!Number.isInteger(maxRebases) || maxRebases < 1 || maxRebases > 10)
		throw new RangeError('Storage merge attempts must be an integer from 1 to 10.');
	let base = clone(opened);
	let acknowledgedLocal = opened.bytes.slice();
	let local = opened.bytes.slice();
	let conflict: StorageConflict<F> | null = null;
	let pending: Promise<unknown> = Promise.resolve();
	function enqueue<T>(run: () => Promise<T>): Promise<T> {
		const result = pending.then(run);
		pending = result.catch(() => {});
		return result;
	}
	function adopt(next: StoredDocument<F>, submitted = next.bytes) {
		base = clone(next);
		acknowledgedLocal = submitted.slice();
		local = next.bytes.slice();
		conflict = null;
	}
	async function save(snapshot: Uint8Array, options?: StorageRequestOptions): Promise<F> {
		local = snapshot.slice();
		conflict = null;
		let ancestor = clone(base);
		let candidate = snapshot.slice();
		let rebases = 0;
		if (merge && !equal(acknowledgedLocal, base.bytes)) {
			// A queued editor snapshot may predate our previous successful remote merge. Rebase its
			// local changes on the actual acknowledged bytes before writing with their fresh eTag.
			conflict = {
				base: { file: cloneFile(base.file), bytes: acknowledgedLocal.slice() },
				local: candidate.slice(),
				remote: clone(base),
			};
			candidate = await mergeConflict(conflict, options);
			local = candidate.slice();
		}
		while (true) {
			options?.signal?.throwIfAborted();
			try {
				const file = await client.save(ancestor.file, candidate, options);
				adopt({ file, bytes: candidate }, snapshot);
				return cloneFile(file);
			} catch (error) {
				if (!client.isConflict(error)) throw error;
			}
			options?.onSaveProgress?.({ phase: 'checking' });
			if (client.latest && client.sameContent) {
				const current = await client.latest(ancestor.file, options);
				if (client.sameContent(ancestor.file, current)) {
					// Only the revision moved (the service touched the file): retry, nothing to merge.
					if (rebases >= maxRebases) {
						conflict = {
							base: clone(ancestor),
							local: candidate.slice(),
							remote: clone(ancestor),
							issues: ['remote-keeps-changing'],
						};
						throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
							cloneConflict(conflict),
						);
					}
					rebases++;
					ancestor = { file: cloneFile(current), bytes: ancestor.bytes.slice() };
					continue;
				}
			}
			const remote = await client.read(ancestor.file, options);
			if (remote.file.kind !== ancestor.file.kind) {
				conflict = {
					base: clone(ancestor),
					local: candidate.slice(),
					remote: clone(remote),
					issues: ['file-format-changed'],
				};
				throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
					cloneConflict(conflict),
				);
			}
			if (equal(candidate, remote.bytes)) {
				adopt(remote, snapshot);
				return cloneFile(remote.file);
			}
			conflict = { base: clone(ancestor), local: candidate.slice(), remote: clone(remote) };
			if (rebases >= maxRebases) {
				conflict.issues = ['remote-keeps-changing'];
				throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
					cloneConflict(conflict),
				);
			}
			rebases++;
			if (!equal(ancestor.bytes, remote.bytes)) {
				if (!merge)
					throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
						cloneConflict(conflict),
					);
				options?.onSaveProgress?.({ phase: 'merging' });
				candidate = await mergeConflict(conflict, options);
				local = candidate.slice();
			}
			ancestor = clone(remote);
		}
	}
	async function mergeConflict(value: StorageConflict<F>, options?: StorageRequestOptions) {
		let result;
		try {
			result = await runStorageMerge(merge!, {
				...cloneConflict(value),
				kind: value.base.file.kind,
				...(options?.signal ? { signal: options.signal } : {}),
			});
		} catch {
			options?.signal?.throwIfAborted();
			value.issues = ['merge-failed'];
			throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
				cloneConflict(value),
			);
		}
		options?.signal?.throwIfAborted();
		if (
			!result ||
			result.status !== 'merged' ||
			!(result.bytes instanceof Uint8Array) ||
			!result.bytes.length
		) {
			value.issues =
				result?.status === 'unresolved' && Array.isArray(result.issues)
					? result.issues.filter((issue): issue is string => typeof issue === 'string')
					: ['invalid-merge-result'];
			throw (settings.conflictError ?? ((value) => new StorageConflictError(value)))(
				cloneConflict(value),
			);
		}
		return result.bytes.slice();
	}
	const session: StorageSession<F> = {
		getSnapshot: () => ({
			file: cloneFile(base.file),
			baseBytes: base.bytes.slice(),
			conflict: conflict ? cloneConflict(conflict) : null,
		}),
		save(bytes, options) {
			const snapshot = bytes.slice();
			return enqueue(() => save(snapshot, options));
		},
		reload(options) {
			return enqueue(async () => {
				const next = await client.read(base.file, options);
				adopt(next);
				return clone(next);
			});
		},
		saveCopy(bytes, name, options) {
			const snapshot = bytes?.slice();
			return enqueue(async () => {
				if (!client.saveCopy)
					throw new Error('This storage provider does not support saving copies.');
				const next = await client.saveCopy(base.file, snapshot ?? local, name, options);
				adopt(next);
				return clone(next);
			});
		},
		storage: {
			async save(bytes, meta) {
				settings.validateSave?.(meta, cloneFile(base.file));
				if (meta.reason === 'saveAs') {
					const result = await session.saveCopy(bytes, meta.fileName);
					return { fileName: result.file.name };
				}
				const result = await session.save(
					bytes,
					meta.onProgress ? { onSaveProgress: meta.onProgress } : undefined,
				);
				return { fileName: result.name };
			},
		},
	};
	return session;
}
