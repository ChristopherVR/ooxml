import { describe, expect, it } from 'vitest';
import { createOneDriveSession, OneDriveConflictError } from './index';
import type { OneDriveMergeInput, OneDriveMergeResult } from './index';
import { createTestDrive } from './__fixtures__/graph';

const encode = (value: Record<string, number>) => new TextEncoder().encode(JSON.stringify(value));
const decode = (value: Uint8Array) =>
	JSON.parse(new TextDecoder().decode(value)) as Record<string, number>;
const combine = (input: OneDriveMergeInput): OneDriveMergeResult => ({
	status: 'merged',
	bytes: encode({ local: decode(input.local).local!, remote: decode(input.remote.bytes).remote! }),
});

describe('OneDrive semantic merge orchestration', () => {
	it('never rebases one Office format onto an item renamed to another format', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]), name: 'Report.pptx' });
		const opened = await drive.client.open(drive.shareUrl);
		const client = {
			...drive.client,
			read: async (...args: Parameters<typeof drive.client.read>) => {
				const remote = await drive.client.read(...args);
				return { ...remote, file: { ...remote.file, kind: 'xlsx' as const, name: 'Renamed.xlsx' } };
			},
		};
		const session = createOneDriveSession(client, opened, {
			merge: () => {
				throw new Error('Must not invoke a mismatched format merger');
			},
		});
		drive.changeRemotely();
		await expect(session.save(new Uint8Array([2]))).rejects.toMatchObject({
			conflict: { issues: ['file-format-changed'] },
		});
		expect(drive.getBytes()).toEqual(new Uint8Array([1]));
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(1);
	});
	it.each(['docx', 'xlsx', 'pptx', 'vsdx'])(
		'saves both changes from the %s merge hook and acknowledges the actual merged bytes',
		async (kind) => {
			const drive = createTestDrive({
				bytes: encode({ local: 0, remote: 0 }),
				name: `Report.${kind}`,
			});
			const seen: OneDriveMergeInput[] = [];
			const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
				merge(input) {
					seen.push(input);
					return combine(input);
				},
			});
			drive.changeRemotely(encode({ local: 0, remote: 2 }));
			const file = await session.save(encode({ local: 1, remote: 0 }));
			expect(seen).toHaveLength(1);
			expect(seen[0]!.kind).toBe(kind);
			expect(decode(seen[0]!.base.bytes)).toEqual({ local: 0, remote: 0 });
			expect(decode(drive.getBytes())).toEqual({ local: 1, remote: 2 });
			expect(session.getSnapshot().file).toEqual(file);
			expect(decode(session.getSnapshot().baseBytes)).toEqual({ local: 1, remote: 2 });
			expect(session.getSnapshot().conflict).toBeNull();
		},
	);

	it('retains an unresolved semantic conflict without attempting a merged cloud upload', async () => {
		const drive = createTestDrive({ bytes: encode({ value: 0 }) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge: () => ({ status: 'unresolved', issues: ['same-cell-changed'] }),
		});
		drive.changeRemotely(encode({ value: 2 }));
		await expect(session.save(encode({ value: 1 }))).rejects.toMatchObject({
			conflict: { issues: ['same-cell-changed'] },
		});
		expect(decode(drive.getBytes())).toEqual({ value: 2 });
		expect(decode(session.getSnapshot().conflict!.local)).toEqual({ value: 1 });
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(1);
	});

	it('sanitizes failed merge hooks and never writes their partial work', async () => {
		const drive = createTestDrive({ bytes: encode({ value: 0 }) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge(input) {
				input.local.fill(9);
				throw new Error('private-provider-details');
			},
		});
		drive.changeRemotely(encode({ value: 2 }));
		const saving = session.save(encode({ value: 1 }));
		await expect(saving).rejects.toMatchObject({ conflict: { issues: ['merge-failed'] } });
		await expect(saving).rejects.not.toThrow('private-provider-details');
		expect(decode(session.getSnapshot().conflict!.local)).toEqual({ value: 1 });
		expect(decode(drive.getBytes())).toEqual({ value: 2 });
	});

	it('rebases on each newer remote ancestor and stops after the configured bound', async () => {
		const drive = createTestDrive({ bytes: encode({ local: 0, remote: 0 }) });
		let writes = 0;
		const ancestors: number[] = [];
		const client = {
			...drive.client,
			save(file: Parameters<typeof drive.client.save>[0], bytes: Uint8Array) {
				drive.changeRemotely(encode({ local: 0, remote: ++writes }));
				return drive.client.save(file, bytes);
			},
		};
		const session = createOneDriveSession(client, await client.open(drive.shareUrl), {
			maxMergeAttempts: 3,
			merge(input) {
				ancestors.push(decode(input.base.bytes).remote!);
				return combine(input);
			},
		});
		await expect(session.save(encode({ local: 1, remote: 0 }))).rejects.toBeInstanceOf(
			OneDriveConflictError,
		);
		expect(writes).toBe(4);
		expect(ancestors).toEqual([0, 1, 2]);
		expect(decode(session.getSnapshot().conflict!.local).local).toBe(1);
		expect(session.getSnapshot().conflict!.issues).toEqual(['remote-keeps-changing']);
		expect(decode(session.getSnapshot().baseBytes)).toEqual({ local: 0, remote: 0 });
		expect(decode(drive.getBytes())).toEqual({ local: 0, remote: 4 });
	});

	it('rechecks access before uploading a merged result', async () => {
		const drive = createTestDrive({ bytes: encode({ local: 0, remote: 0 }) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge(input) {
				drive.setAccess('read');
				return combine(input);
			},
		});
		drive.changeRemotely(encode({ local: 0, remote: 2 }));
		await expect(session.save(encode({ local: 1, remote: 0 }))).rejects.toMatchObject({
			code: 'access-denied',
		});
		expect(decode(drive.getBytes())).toEqual({ local: 0, remote: 2 });
		expect(decode(session.getSnapshot().baseBytes)).toEqual({ local: 0, remote: 0 });
	});

	it('cancels a pending asynchronous merge without uploading when it later resolves', async () => {
		const drive = createTestDrive({ bytes: encode({ local: 0, remote: 0 }) });
		let finish: ((value: OneDriveMergeResult) => void) | undefined;
		let entered: (() => void) | undefined;
		const started = new Promise<void>((resolve) => {
			entered = resolve;
		});
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge() {
				entered!();
				return new Promise<OneDriveMergeResult>((resolve) => {
					finish = resolve;
				});
			},
		});
		drive.changeRemotely(encode({ local: 0, remote: 2 }));
		const abort = new AbortController();
		const saving = session.save(encode({ local: 1, remote: 0 }), { signal: abort.signal });
		await started;
		abort.abort();
		await expect(saving).rejects.toMatchObject({ name: 'AbortError' });
		finish!({ status: 'merged', bytes: encode({ local: 1, remote: 2 }) });
		await Promise.resolve();
		expect(decode(drive.getBytes())).toEqual({ local: 0, remote: 2 });
		expect(
			drive.requests.filter((request) => request.url.endsWith('/createUploadSession')),
		).toHaveLength(1);
	});

	it('serializes later local saves using the previously acknowledged merged ancestor', async () => {
		const drive = createTestDrive({ bytes: encode({ local: 0, remote: 0 }) });
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge: combine,
		});
		drive.changeRemotely(encode({ local: 0, remote: 2 }));
		await session.save(encode({ local: 1, remote: 0 }));
		const merged = decode(session.getSnapshot().baseBytes);
		await session.save(encode({ ...merged, local: 3 }));
		expect(decode(drive.getBytes())).toEqual({ local: 3, remote: 2 });
	});

	it('rebases an already queued local snapshot onto the earlier save’s merged bytes', async () => {
		const drive = createTestDrive({ bytes: encode({ local: 0, remote: 0 }) });
		let entered: (() => void) | undefined;
		let release: (() => void) | undefined;
		const started = new Promise<void>((resolve) => {
			entered = resolve;
		});
		const wait = new Promise<void>((resolve) => {
			release = resolve;
		});
		const ancestors: Array<Record<string, number>> = [];
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			async merge(input) {
				ancestors.push(decode(input.base.bytes));
				if (ancestors.length === 1) {
					entered!();
					await wait;
				}
				return combine(input);
			},
		});
		drive.changeRemotely(encode({ local: 0, remote: 2 }));
		const first = session.save(encode({ local: 1, remote: 0 }));
		await started;
		const later = session.save(encode({ local: 3, remote: 0 }));
		release!();
		await Promise.all([first, later]);
		expect(ancestors).toEqual([
			{ local: 0, remote: 0 },
			{ local: 1, remote: 0 },
		]);
		expect(decode(drive.getBytes())).toEqual({ local: 3, remote: 2 });
		expect(decode(session.getSnapshot().baseBytes)).toEqual({ local: 3, remote: 2 });
	});

	it.each([0, -1, 11, 1.5, Number.NaN])('rejects invalid merge retry bound %s', async (limit) => {
		const drive = createTestDrive({ bytes: encode({ value: 0 }) });
		const opened = await drive.client.open(drive.shareUrl);
		expect(() =>
			createOneDriveSession(drive.client, opened, { merge: combine, maxMergeAttempts: limit }),
		).toThrow(RangeError);
	});
});
