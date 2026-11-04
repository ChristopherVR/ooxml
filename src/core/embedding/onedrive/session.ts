import { officeKind } from './metadata';
import { OneDriveConflictError, OneDriveError } from './types';
import type {
	OneDriveClient,
	OneDriveConflict,
	OneDriveFile,
	OneDriveOpened,
	OneDriveRequestOptions,
	OneDriveSession,
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
});

/**
 * One serialized cloud document identity and its last acknowledged baseline. Bytes from failed
 * saves remain available for conflict recovery. This is storage coordination, not Office merging.
 */
export function createOneDriveSession(
	client: OneDriveClient,
	opened: OneDriveOpened,
): OneDriveSession {
	let base = clone(opened);
	let local = opened.bytes.slice();
	let conflict: OneDriveConflict | null = null;
	let pending: Promise<unknown> = Promise.resolve();
	function enqueue<T>(run: () => Promise<T>): Promise<T> {
		const result = pending.then(run);
		pending = result.catch(() => {});
		return result;
	}
	function adopt(next: OneDriveOpened) {
		base = clone(next);
		local = next.bytes.slice();
		conflict = null;
	}
	async function save(
		snapshot: Uint8Array,
		options?: OneDriveRequestOptions,
	): Promise<OneDriveFile> {
		local = snapshot.slice();
		try {
			const file = await client.save(base.file, snapshot, options);
			adopt({ file, bytes: snapshot });
			return { ...file };
		} catch (error) {
			if (!(error instanceof OneDriveError) || error.code !== 'conflict') throw error;
			const remote = await client.read(base.file, options);
			if (equal(snapshot, remote.bytes)) {
				// The previous response may have been lost after a successful commit.
				adopt(remote);
				return { ...remote.file };
			}
			if (equal(base.bytes, remote.bytes)) {
				// A revision changed without changing content: retry once against the fresh revision.
				try {
					const file = await client.save(remote.file, snapshot, options);
					adopt({ file, bytes: snapshot });
					return { ...file };
				} catch (retryError) {
					if (!(retryError instanceof OneDriveError) || retryError.code !== 'conflict')
						throw retryError;
					const latest = await client.read(base.file, options);
					conflict = { base: clone(base), local: snapshot.slice(), remote: clone(latest) };
				}
			} else conflict = { base: clone(base), local: snapshot.slice(), remote: clone(remote) };
			throw new OneDriveConflictError(cloneConflict(conflict));
		}
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
