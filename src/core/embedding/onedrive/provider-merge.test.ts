import { describe, expect, it, vi } from 'vitest';
import { createOneDriveFileProvider } from './index';
import type { OneDriveFileProviderOptions } from './index';
import { createTestDrive } from './__fixtures__/graph';

const merged = new Uint8Array([9]);

async function saveOverRemoteChange(options: Partial<OneDriveFileProviderOptions>) {
	const drive = createTestDrive({ bytes: new Uint8Array([1]), name: 'Deck.pptx' });
	const provider = createOneDriveFileProvider({
		fetch: drive.fetch,
		getAccessToken: async () => 'test-token',
		canWrite: () => true,
		account: { name: 'Ada' },
		signIn: async () => ({ name: 'Ada' }),
		...options,
	});
	const opened = await provider.openLink!(drive.shareUrl);
	drive.changeRemotely(new Uint8Array([2]));
	const saved = opened.storage.save(new Uint8Array([3]), {
		fileName: 'Deck.pptx',
		mimeType: 'application/octet-stream',
		reason: 'save',
	});
	return { drive, saved };
}

describe('OneDrive file provider merge configuration', () => {
	it('uses a host merge for every kind in place of the per-kind mergers', async () => {
		const perKind = vi.fn();
		const merge = vi.fn(() => ({ status: 'merged' as const, bytes: merged }));
		const { drive, saved } = await saveOverRemoteChange({ merge, mergers: { pptx: perKind } });
		await saved;
		expect(merge).toHaveBeenCalledOnce();
		expect(perKind).not.toHaveBeenCalled();
		expect(drive.getBytes()).toEqual(merged);
	});

	it('never merges when merge is false, and reports the conflict', async () => {
		const perKind = vi.fn(() => ({ status: 'merged' as const, bytes: merged }));
		const { drive, saved } = await saveOverRemoteChange({
			merge: false,
			mergers: { pptx: perKind },
		});
		await expect(saved).rejects.toMatchObject({ name: 'OneDriveConflictError' });
		expect(perKind).not.toHaveBeenCalled();
		expect(drive.getBytes()).toEqual(new Uint8Array([2]));
	});
});
