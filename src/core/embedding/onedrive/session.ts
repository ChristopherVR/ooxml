import { officeKind } from './metadata';
import { runOneDriveMerge } from './merge';
import { OneDriveConflictError, OneDriveError } from './types';
import type {
	OneDriveClient,
	OneDriveConflict,
	OneDriveFile,
	OneDriveOpened,
	OneDriveRequestOptions,
	OneDriveSession,
	OneDriveSessionOptions,
} from './types';

const clone = (value: OneDriveOpened): OneDriveOpened => ({
	file: { ...value.file },
	bytes: value.bytes.slice(),
});
const equal = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((value, index) => value === b[index]);
const cloneConflict = (value: OneDriveConflict): OneDriveConflict => ({
	base: clone(value.base),
	local: value.local.slice(),
	remote: clone(value.remote),
	...(value.issues ? { issues: [...value.issues] } : {}),
});

/**
 * One serialized cloud document identity and its last acknowledged baseline. Bytes from failed
 * saves remain available for conflict recovery. This is storage coordination, not Office merging.
 */
export function createOneDriveSession(
	client: OneDriveClient,
	opened: OneDriveOpened,
	settings: OneDriveSessionOptions = {},
): OneDriveSession {
	const merge = settings.merge;
	const maxRebases = merge ? (settings.maxMergeAttempts ?? 3) : 1;
	if (!Number.isInteger(maxRebases) || maxRebases < 1 || maxRebases > 10)
		throw new RangeError('OneDrive merge attempts must be an integer from 1 to 10.');
	let base = clone(opened);
	let acknowledgedLocal = opened.bytes.slice();
	let local = opened.bytes.slice();
	let conflict: OneDriveConflict | null = null;
	let pending: Promise<unknown> = Promise.resolve();
	function enqueue<T>(run: () => Promise<T>): Promise<T> {
		const result = pending.then(run);
		pending = result.catch(() => {});
		return result;
	}
	function adopt(next: OneDriveOpened, submitted = next.bytes) {
		base = clone(next);
		acknowledgedLocal = submitted.slice();
		local = next.bytes.slice();
		conflict = null;
	}
	async function save(
		snapshot: Uint8Array,
		options?: OneDriveRequestOptions,
	): Promise<OneDriveFile> {
		local = snapshot.slice();
		conflict = null;
		let ancestor = clone(base);
		let candidate = snapshot.slice();
		let rebases = 0;
		if (merge && !equal(acknowledgedLocal, base.bytes)) {
			// A queued editor snapshot may predate our previous successful remote merge. Rebase its
			// local changes on the actual acknowledged bytes before writing with their fresh eTag.
			conflict = {
				base: { file: { ...base.file }, bytes: acknowledgedLocal.slice() },
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
				return { ...file };
			} catch (error) {
				if (!(error instanceof OneDriveError) || error.code !== 'conflict') throw error;
			}
			const remote = await client.read(ancestor.file, options);
			if (remote.file.kind !== ancestor.file.kind) {
				conflict = {
					base: clone(ancestor),
					local: candidate.slice(),
					remote: clone(remote),
					issues: ['file-format-changed'],
				};
				throw new OneDriveConflictError(cloneConflict(conflict));
			}
			if (equal(candidate, remote.bytes)) {
				adopt(remote, snapshot);
				return { ...remote.file };
			}
			conflict = { base: clone(ancestor), local: candidate.slice(), remote: clone(remote) };
			if (rebases >= maxRebases) {
				conflict.issues = ['remote-keeps-changing'];
				throw new OneDriveConflictError(cloneConflict(conflict));
			}
			rebases++;
			if (!equal(ancestor.bytes, remote.bytes)) {
				if (!merge) throw new OneDriveConflictError(cloneConflict(conflict));
				candidate = await mergeConflict(conflict, options);
				local = candidate.slice();
			}
			ancestor = clone(remote);
		}
	}
	async function mergeConflict(value: OneDriveConflict, options?: OneDriveRequestOptions) {
		let result;
		try {
			result = await runOneDriveMerge(merge!, {
				...cloneConflict(value),
				kind: value.base.file.kind,
				...(options?.signal ? { signal: options.signal } : {}),
			});
		} catch {
			options?.signal?.throwIfAborted();
			value.issues = ['merge-failed'];
			throw new OneDriveConflictError(cloneConflict(value));
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
			throw new OneDriveConflictError(cloneConflict(value));
		}
		return result.bytes.slice();
	}
	const session: OneDriveSession = {
		getSnapshot: () => ({
			file: { ...base.file },
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
				const next = await client.saveCopy(base.file, snapshot ?? local, name, options);
				adopt(next);
				return clone(next);
			});
		},
		storage: {
			async save(bytes, meta) {
				if (officeKind(meta.fileName) !== base.file.kind)
					throw new OneDriveError(
						'unsupported',
						'A different Office format needs a separate export destination.',
					);
				if (meta.reason === 'saveAs') {
					const result = await session.saveCopy(bytes, meta.fileName);
					return { fileName: result.file.name };
				}
				const result = await session.save(bytes);
				return { fileName: result.name };
			},
		},
	};
	return session;
}
