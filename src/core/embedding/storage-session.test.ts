import { describe, expect, it, vi } from 'vitest';
import { createStorageSession } from './storage-session';
import { createStorageMerger } from './storage-merge';
import { StorageConflictError } from './versioned-storage';
import type { StorageMerger, VersionedStorageProvider } from './versioned-storage';

const bytes = (value: string) => new TextEncoder().encode(value);
const text = (value: Uint8Array) => new TextDecoder().decode(value);
interface File {
	name: string;
	kind: string;
	key: string;
	revision: number;
	nested: { owner: string };
}
function setup() {
	let current = {
		file: {
			name: 'Report.docx',
			kind: 'docx',
			key: '/report',
			revision: 1,
			nested: { owner: 'host' },
		},
		bytes: bytes('base'),
	};
	const stale = new Error('stale');
	const provider: VersionedStorageProvider<File> = {
		read: vi.fn(async () => structuredClone(current)),
		save: vi.fn(async (file, data, options) => {
			options?.signal?.throwIfAborted();
			if (file.revision !== current.file.revision) throw stale;
			current = { file: { ...file, revision: file.revision + 1 }, bytes: data.slice() };
			return structuredClone(current.file);
		}),
		isConflict: (error) => error === stale,
	};
	return {
		provider,
		opened: structuredClone(current),
		remote(value = 'remote') {
			current = {
				file: { ...current.file, revision: current.file.revision + 1 },
				bytes: bytes(value),
			};
		},
		current: () => structuredClone(current),
	};
}
const merge: StorageMerger<File> = ({ local, remote }) => ({
	status: 'merged',
	bytes: bytes(`${text(local)}+${text(remote.bytes)}`),
});

describe('provider-neutral storage sessions', () => {
	it('conditionally saves and rebases with host-owned identities and numeric revisions', async () => {
		const drive = setup();
		const session = createStorageSession(drive.provider, drive.opened, { merge });
		drive.remote();
		const result = await session.save(bytes('local'));
		expect(result.revision).toBe(3);
		expect(text(session.getSnapshot().baseBytes)).toBe('local+remote');
		expect(text(drive.current().bytes)).toBe('local+remote');
	});
	it('retains three versions without a merger and never overwrites the latest', async () => {
		const drive = setup();
		const session = createStorageSession(drive.provider, drive.opened);
		drive.remote();
		await expect(session.save(bytes('local'))).rejects.toBeInstanceOf(StorageConflictError);
		expect(text(session.getSnapshot().conflict!.local)).toBe('local');
		expect(text(drive.current().bytes)).toBe('remote');
	});
	it('isolates nested metadata and input buffers from external mutation', async () => {
		const drive = setup();
		const session = createStorageSession(drive.provider, drive.opened);
		drive.opened.file.nested.owner = 'changed';
		const snapshot = session.getSnapshot();
		snapshot.file.nested.owner = 'also changed';
		snapshot.baseBytes.fill(0);
		expect(session.getSnapshot().file.nested.owner).toBe('host');
		expect(text(session.getSnapshot().baseBytes)).toBe('base');
	});
	it('propagates authentication failures without attempting a merge', async () => {
		const drive = setup();
		const failure = new Error('sign-in required');
		drive.provider.save = vi.fn(async () => {
			throw failure;
		});
		const combine = vi.fn(merge);
		const session = createStorageSession(drive.provider, drive.opened, { merge: combine });
		await expect(session.save(bytes('local'))).rejects.toBe(failure);
		expect(combine).not.toHaveBeenCalled();
		expect(drive.provider.read).not.toHaveBeenCalled();
	});
	it('allows custom validation and errors without invoking the provider', async () => {
		const drive = setup();
		const invalid = new Error('use an export destination');
		const session = createStorageSession(drive.provider, drive.opened, {
			validateSave: () => {
				throw invalid;
			},
		});
		await expect(
			session.storage.save(bytes('export'), {
				fileName: 'Report.pdf',
				mimeType: 'application/pdf',
				reason: 'save',
			}),
		).rejects.toBe(invalid);
		expect(drive.provider.save).not.toHaveBeenCalled();
	});
	it('does not fake Save As when the provider cannot create a copy', async () => {
		const drive = setup();
		const session = createStorageSession(drive.provider, drive.opened);
		await expect(session.saveCopy(bytes('local'))).rejects.toThrow('does not support');
		expect(drive.provider.save).not.toHaveBeenCalled();
	});
	it('adopts a provider-created copy as a new identity and supports host conflict errors', async () => {
		const drive = setup();
		drive.provider.saveCopy = vi.fn(async (file, data, name) => ({
			file: { ...file, key: '/copy', name: name ?? file.name, revision: 1 },
			bytes: data.slice(),
		}));
		const customError = new Error('host conflict dialog');
		const session = createStorageSession(drive.provider, drive.opened, {
			conflictError: () => customError,
		});
		drive.remote();
		await expect(session.save(bytes('local'))).rejects.toBe(customError);
		expect(session.getSnapshot().conflict).not.toBeNull();
		await session.saveCopy(undefined, 'Copy.docx');
		expect(session.getSnapshot().file.key).toBe('/copy');
		expect(session.getSnapshot().file.name).toBe('Copy.docx');
		expect(session.getSnapshot().conflict).toBeNull();
		expect(text(session.getSnapshot().baseBytes)).toBe('local');
		expect(text(drive.current().bytes)).toBe('remote');
	});
	it('stops retrying when another writer keeps changing the file', async () => {
		const drive = setup();
		const session = createStorageSession(drive.provider, drive.opened, {
			maxMergeAttempts: 2,
			merge: (input) => {
				drive.remote('changed again');
				return merge(input);
			},
		});
		drive.remote();
		await expect(session.save(bytes('local'))).rejects.toBeInstanceOf(StorageConflictError);
		expect(session.getSnapshot().conflict?.issues).toEqual(['remote-keeps-changing']);
		expect(drive.provider.save).toHaveBeenCalledTimes(3);
	});
});

describe('custom merge routing', () => {
	const input = () => {
		const { opened } = setup();
		return { kind: 'docx', base: opened, local: bytes('local'), remote: opened };
	};
	it('routes host formats, supports overrides and explicitly disabled handlers', async () => {
		const custom = vi.fn(merge);
		const router = createStorageMerger<File>({ docx: custom, xlsx: undefined });
		expect((await router(input())).status).toBe('merged');
		expect(custom).toHaveBeenCalledOnce();
		expect(await router({ ...input(), kind: 'xlsx' })).toEqual({
			status: 'unresolved',
			issues: ['unsupported-merge-kind:xlsx'],
		});
		expect((await router({ ...input(), kind: 'toString' })).status).toBe('unresolved');
	});
	it('runs an optional fallback and cancels an uncooperative host merger', async () => {
		const fallback = vi.fn(merge);
		await createStorageMerger<File>({}, { fallback })(input());
		expect(fallback).toHaveBeenCalledOnce();
		expect(
			(await createStorageMerger<File>({ docx: undefined }, { fallback })(input())).status,
		).toBe('unresolved');
		expect(fallback).toHaveBeenCalledOnce();
		const controller = new AbortController();
		const router = createStorageMerger<File>({ docx: () => new Promise(() => {}) });
		const pending = router({ ...input(), signal: controller.signal });
		controller.abort(new Error('canceled'));
		await expect(pending).rejects.toThrow('canceled');
	});
});
