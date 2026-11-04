import { describe, expect, it, vi } from 'vitest';
import { createOneDriveClient, createOneDriveFileProvider } from './index';
import { encodeEntryId } from './browse';

const bytes = new Uint8Array([80, 75, 3, 4, 1, 2, 3, 4]);
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
const meta = {
	id: 'item',
	name: 'Deck.pptx',
	eTag: '"v1"',
	size: bytes.length,
	parentReference: { driveId: 'drive' },
	file: {},
	'@microsoft.graph.downloadUrl': 'https://download.example/f?secret=1',
};
const link = 'https://1drv.ms/p/x';
const isDownload = (url: unknown) => String(url).startsWith('https://download');
const entry = {
	id: encodeEntryId({ driveId: 'drive', itemId: 'item' }),
	name: 'Deck.pptx',
	type: 'file' as const,
};

describe('OneDrive file provider', () => {
	const make = (extra = {}) =>
		createOneDriveFileProvider({
			getAccessToken: async () => 't',
			signIn: async () => ({ name: 'Ada' }),
			fetch: async (url) => (isDownload(url) ? new Response(bytes) : json(meta)),
			...extra,
		});

	it('rejects with an authentication error until signed in, and notifies', async () => {
		const provider = make();
		for (const call of [
			() => provider.list(),
			() => provider.recent?.(),
			() => provider.search?.('x'),
			() => provider.open(entry),
			() => provider.openLink?.(link),
		])
			await expect(call()).rejects.toMatchObject({ code: 'authentication' });
		const listener = vi.fn();
		const stop = provider.subscribe?.(listener);
		await provider.signIn();
		expect(provider.account()).toEqual({ name: 'Ada' });
		await provider.signOut?.();
		expect(provider.account()).toBeNull();
		expect(listener).toHaveBeenCalledTimes(2);
		stop?.();
		await provider.signIn();
		expect(listener).toHaveBeenCalledTimes(2);
	});

	it('opens by identity with progress and reports canEdit from the host policy', async () => {
		const provider = make({ account: { name: 'Ada' }, canWrite: () => true });
		const progress: number[] = [];
		const opened = await provider.open(entry, { onProgress: (loaded) => progress.push(loaded) });
		expect(opened).toMatchObject({ kind: 'pptx', fileName: 'Deck.pptx', canEdit: true });
		expect(opened.bytes).toEqual(bytes);
		expect(progress.at(-1)).toBe(bytes.length);
		expect(typeof opened.storage.save).toBe('function');
		const readOnly = make({ account: { name: 'Ada' } });
		expect((await readOnly.openLink?.(link))?.canEdit).toBe(false);
	});
});

describe('OneDrive download speed', () => {
	it('starts the download before the write policy resolves, with one metadata request', async () => {
		const order: string[] = [];
		let release!: (value: boolean) => void;
		const policy = new Promise<boolean>((resolve) => (release = resolve));
		const client = createOneDriveClient({
			getAccessToken: async () => 't',
			canWrite: () => (order.push('canWrite'), policy),
			fetch: async (url) => {
				order.push(isDownload(url) ? 'download' : 'graph');
				return isDownload(url) ? new Response(bytes) : json(meta);
			},
		});
		const opening = client.openItem({ driveId: 'drive', itemId: 'item' });
		await vi.waitFor(() => expect(order).toEqual(['graph', 'canWrite', 'download']));
		release(true);
		expect((await opening).file.access).toBe('write');
		expect(order.filter((step) => step === 'graph')).toHaveLength(1);
	});

	it('sends $select and Prefer: redeemSharingLink on the shares request, token once', async () => {
		const calls: Array<{ url: string; prefer: string | null }> = [];
		const getAccessToken = vi.fn(async () => 't');
		const client = createOneDriveClient({
			getAccessToken,
			fetch: async (url, init) => {
				calls.push({ url: String(url), prefer: new Headers(init?.headers).get('Prefer') });
				return isDownload(url) ? new Response(bytes) : json(meta);
			},
		});
		await client.open(link);
		expect(calls[0]?.url).toMatch(/\/shares\/u![^/]+\/driveItem\?\$select=.*downloadUrl$/u);
		expect(calls[0]?.prefer).toBe('redeemSharingLink');
		expect(calls[1]?.prefer).toBeNull();
		expect(getAccessToken).toHaveBeenCalledTimes(1);
	});

	it('reports progress per chunk and preallocates from Content-Length', async () => {
		const stream = new ReadableStream<Uint8Array>({
			start(controller) {
				controller.enqueue(bytes.subarray(0, 3));
				controller.enqueue(bytes.subarray(3));
				controller.close();
			},
		});
		const client = createOneDriveClient({
			getAccessToken: async () => 't',
			fetch: async (url) =>
				isDownload(url)
					? new Response(stream, { headers: { 'Content-Length': String(bytes.length) } })
					: json(meta),
		});
		const seen: Array<[number, number | undefined]> = [];
		const opened = await client.openItem(
			{ driveId: 'drive', itemId: 'item' },
			{ onProgress: (loaded, total) => seen.push([loaded, total]) },
		);
		expect(seen).toEqual([
			[3, bytes.length],
			[bytes.length, bytes.length],
		]);
		expect(opened.bytes).toEqual(bytes);
	});

	it('reads metadata with one request and no download', async () => {
		const urls: string[] = [];
		const client = createOneDriveClient({
			getAccessToken: async () => 't',
			fetch: async (url) => (urls.push(String(url)), json(meta)),
		});
		const current = await client.metadata({
			driveId: 'drive',
			itemId: 'item',
			name: 'Deck.pptx',
			kind: 'pptx',
			eTag: '"v0"',
			access: 'unknown',
		});
		expect(current.eTag).toBe('"v1"');
		expect(urls).toHaveLength(1);
	});
});
