import { describe, expect, it, vi } from 'vitest';
import type { DocumentSaveProgress } from '../storage';
import { createOneDriveSession } from './index';
import { createTestDrive } from './__fixtures__/graph';

const downloads = (drive: ReturnType<typeof createTestDrive>) =>
	drive.requests.filter((request) => request.url.includes('/download/')).length;
const meta = (onProgress?: (progress: DocumentSaveProgress) => void) => ({
	fileName: 'Deck.pptx',
	mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
	reason: 'save' as const,
	...(onProgress ? { onProgress } : {}),
});

describe('OneDrive save progress and metadata-only changes', () => {
	it('saves twice in one session with no conflict and no download', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const merge = vi.fn();
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge,
		});
		const before = downloads(drive);
		await session.storage.save(new Uint8Array([2]), meta());
		await session.storage.save(new Uint8Array([3]), meta());
		expect(drive.getBytes()).toEqual(new Uint8Array([3]));
		expect(downloads(drive)).toBe(before);
		expect(merge).not.toHaveBeenCalled();
	});

	it('retries over a metadata-only change without downloading or merging', async () => {
		const drive = createTestDrive({ bytes: new Uint8Array([1]) });
		const merge = vi.fn();
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge,
		});
		await session.storage.save(new Uint8Array([2]), meta());
		drive.touch(); // OneDrive changed the eTag after the upload, not the content
		const before = downloads(drive);
		const progress: DocumentSaveProgress[] = [];
		await session.storage.save(
			new Uint8Array([3]),
			meta((step) => progress.push(step)),
		);
		expect(drive.getBytes()).toEqual(new Uint8Array([3]));
		expect(downloads(drive)).toBe(before);
		expect(merge).not.toHaveBeenCalled();
		expect(progress.map((step) => step.phase)).toContain('checking');
		expect(progress.map((step) => step.phase)).not.toContain('downloading');
	});

	it('reports upload progress per fragment, then checking, downloading and merging on a real change', async () => {
		const size = 6 * 1024 * 1024; // two 5 MiB fragments
		const drive = createTestDrive({
			bytes: new Uint8Array(size),
			uploadFragmentBytes: 5 * 1024 * 1024,
		});
		const session = createOneDriveSession(drive.client, await drive.client.open(drive.shareUrl), {
			merge: ({ local }) => ({ status: 'merged', bytes: local }),
		});
		drive.changeRemotely(new Uint8Array(size).fill(7));
		const progress: DocumentSaveProgress[] = [];
		await session.storage.save(
			new Uint8Array(size).fill(1),
			meta((step) => progress.push(step)),
		);
		const phases = progress.map((step) => step.phase);
		expect(phases.indexOf('checking')).toBeGreaterThan(-1);
		expect(phases.indexOf('downloading')).toBeGreaterThan(phases.indexOf('checking'));
		expect(phases.indexOf('merging')).toBeGreaterThan(phases.indexOf('downloading'));
		const uploads = progress.filter((step) => step.phase === 'uploading');
		expect(uploads.at(-1)).toEqual({ phase: 'uploading', loaded: size, total: size });
		expect(uploads.some((step) => step.loaded > 0 && step.loaded < size)).toBe(true);
	});
});
