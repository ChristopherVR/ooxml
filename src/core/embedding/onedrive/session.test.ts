import { describe, expect, it } from 'vitest';
import { createOneDriveSession, OneDriveConflictError, OneDriveError } from './index';
import { createTestDrive } from './__fixtures__/graph';

describe('OneDrive document sessions', () => {
	it.each(['docx', 'xlsx', 'pptx', 'vsdx'])(
		'keeps %s read-only when host access is revoked',
		async (kind) => {
			const drive = createTestDrive({ bytes: new Uint8Array([1]), name: `Report.${kind}` });
			const opened = await drive.client.open(drive.shareUrl);
			drive.setAccess('read');
			expect((await drive.client.read(opened.file)).file.access).toBe('read');
			const session = createOneDriveSession(drive.client, opened);
			await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({
				code: 'access-denied',
			});
			await expect(session.saveCopy(new Uint8Array([2]))).rejects.toMatchObject({
				code: 'access-denied',
			});
			await expect(
				drive.client.share(opened.file, { type: 'view', scope: 'organization' }),
			).rejects.toMatchObject({ code: 'access-denied' });
			expect(drive.getBytes()).toEqual(new Uint8Array([1]));
		},
	);

	it('retains the baseline after an uncertain upload response and recognizes that commit on retry', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		let loseResponse = true;
		const session = createOneDriveSession(
			{
				...drive.client,
				async save(file, bytes, options) {
					const updated = await drive.client.save(file, bytes, options);
					if (loseResponse) {
						loseResponse = false;
						throw new OneDriveError('network', 'Connection interrupted.');
					}
					return updated;
				},
			},
			await drive.client.open(drive.shareUrl),
		);
		await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({ code: 'network' });
		expect(session.getSnapshot().baseBytes).toEqual(new Uint8Array([1]));
		expect(drive.getBytes()).toEqual(new Uint8Array([2]));
		await session.save(new Uint8Array([2]));
		expect(session.getSnapshot().baseBytes).toEqual(new Uint8Array([2]));
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(2);
	});
	it('does not retry indefinitely when the remote file changes again during revision recovery', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const opened = await drive.client.open(drive.shareUrl);
		let attempts = 0;
		const session = createOneDriveSession(
			{
				...drive.client,
				save(file, bytes, options) {
					if (++attempts === 2) drive.changeRemotely(new Uint8Array([3]));
					return drive.client.save(file, bytes, options);
				},
			},
			opened,
		);
		drive.changeRemotely();
		await expect(session.save(new Uint8Array([2]))).rejects.toBeInstanceOf(OneDriveConflictError);
		expect(attempts).toBe(2);
		expect(drive.getBytes()).toEqual(new Uint8Array([3]));
		expect(session.getSnapshot().conflict!.remote.bytes).toEqual(new Uint8Array([3]));
	});

	it('does not reinterpret a server lock as a harmless revision change', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		let reads = 0;
		const session = createOneDriveSession(
			{
				...drive.client,
				async save() {
					throw new OneDriveError('locked', 'Locked', 423);
				},
				async read(file, options) {
					reads++;
					return drive.client.read(file, options);
				},
			},
			await drive.client.open(drive.shareUrl),
		);
		await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({ code: 'locked' });
		expect(reads).toBe(0);
		expect(session.getSnapshot().baseBytes).toEqual(new Uint8Array([1]));
	});

	it('keeps the conflict and original identity when saving a separate copy fails', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		drive.changeRemotely(new Uint8Array([3]));
		await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({ code: 'conflict' });
		drive.setAccess('read');
		await expect(session.saveCopy()).rejects.toMatchObject({ code: 'access-denied' });
		expect(session.getSnapshot().file.itemId).toBe('item');
		expect(session.getSnapshot().conflict!.local).toEqual(new Uint8Array([2]));
	});

	it.each(['../escape.pptx', 'Changed.xlsx', 'folder/name.pptx'])(
		'rejects unsafe or converted copy name %s before uploading',
		async (name) => {
			const drive = createTestDrive({ bytes: new Uint8Array([1]) });
			const opened = await drive.client.open(drive.shareUrl);
			await expect(async () =>
				drive.client.saveCopy(opened.file, opened.bytes, name),
			).rejects.toMatchObject({ code: 'unsupported' });
			expect(drive.requests.some((request) => request.url.endsWith('/createUploadSession'))).toBe(
				false,
			);
		},
	);
	it.each(['docx', 'xlsx', 'pptx', 'vsdx'])(
		'opens, saves, copies and shares %s through the shared transport',
		async (kind) => {
			const drive = createTestDrive({ bytes: new Uint8Array([1]), name: `Report.${kind}` });
			const opened = await drive.client.open(drive.shareUrl);
			expect(opened.file.kind).toBe(kind);
			const session = createOneDriveSession(drive.client, opened);
			await session.storage.save(new Uint8Array([2]), {
				fileName: `Report.${kind}`,
				mimeType: 'application/octet-stream',
				reason: 'save',
			});
			expect(drive.getBytes()).toEqual(new Uint8Array([2]));
			await session.storage.save(new Uint8Array([3]), {
				fileName: `Copy.${kind}`,
				mimeType: 'application/octet-stream',
				reason: 'saveAs',
			});
			const copy = session.getSnapshot().file;
			expect(copy.itemId).not.toBe(opened.file.itemId);
			expect(drive.getBytes()).toEqual(new Uint8Array([2]));
			expect((await session.reload()).bytes).toEqual(new Uint8Array([3]));
			await drive.client.share(copy, { type: 'view', scope: 'organization' });
			const request = drive.requests.at(-1)!;
			expect(JSON.parse(String(request.init?.body))).toEqual({
				type: 'view',
				scope: 'organization',
			});
		},
	);

	it('refreshes a metadata-only revision once and saves without asking the user to discard work', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		drive.changeRemotely();
		await session.save(new Uint8Array([2]));
		expect(drive.getBytes()).toEqual(new Uint8Array([2]));
		expect(session.getSnapshot().conflict).toBeNull();
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(2);
	});

	it('retains base, local and remote bytes on a genuine conflict; cloud copy leaves original intact', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		drive.changeRemotely(new Uint8Array([3]));
		await expect(session.save(new Uint8Array([2]))).rejects.toBeInstanceOf(OneDriveConflictError);
		const state = session.getSnapshot();
		expect(state.conflict).toMatchObject({
			base: { bytes: new Uint8Array([1]) },
			local: new Uint8Array([2]),
			remote: { bytes: new Uint8Array([3]) },
		});
		state.conflict!.local.fill(9);
		expect(session.getSnapshot().conflict!.local).toEqual(new Uint8Array([2]));
		const copied = await session.saveCopy();
		expect(copied.file.name).toBe('Deck (copy).pptx');
		expect(copied.bytes).toEqual(new Uint8Array([2]));
		expect(drive.getBytes()).toEqual(new Uint8Array([3]));
		expect(session.getSnapshot().conflict).toBeNull();
	});

	it('explicit reload adopts the latest version after a conflict', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		drive.changeRemotely(new Uint8Array([3]));
		await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({ code: 'conflict' });
		expect((await session.reload()).bytes).toEqual(new Uint8Array([3]));
		expect(session.getSnapshot().conflict).toBeNull();
	});

	it('recognizes an already committed save after its response was lost', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		drive.changeRemotely(new Uint8Array([2]));
		await session.save(new Uint8Array([2]));
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(1);
		expect(session.getSnapshot().baseBytes).toEqual(new Uint8Array([2]));
	});

	it('serializes simultaneous saves while snapshotting each caller input', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl));
		const first = new Uint8Array([2]);
		const second = new Uint8Array([3]);
		const saves = [session.save(first), session.save(second)];
		first.fill(8);
		second.fill(9);
		await Promise.all(saves);
		expect(drive.getBytes()).toEqual(new Uint8Array([3]));
		expect(session.getSnapshot().baseBytes).toEqual(new Uint8Array([3]));
	});

	it('resumes from persisted identity without web URL and rejects format conversion', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const opened = await drive.client.open(drive.shareUrl);
		delete opened.file.webUrl;
		const session = createOneDriveSession(drive.client, opened);
		await session.save(new Uint8Array([2]));
		await expect(
			session.storage.save(new Uint8Array([3]), {
				fileName: 'Other.xlsx',
				mimeType: '',
				reason: 'saveAs',
			}),
		).rejects.toMatchObject({ code: 'unsupported' });
		expect(drive.getBytes()).toEqual(new Uint8Array([2]));
	});
});
