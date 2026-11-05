import { createOneDriveClient } from '../client';

/** Test-only HTTP fixture. Excluded from declarations and every published bundle. */
export function createTestDrive(options: { bytes: Uint8Array; name?: string; selectDropsDownloadUrl?: boolean }) {
	type Item = { id: string; bytes: Uint8Array; name: string; revision: number };
	const first: Item = { id: 'item', bytes: options.bytes.slice(), name: options.name ?? 'Deck.pptx', revision: 1 };
	const items = new Map<string, Item>([['item', first]]);
	const uploads = new Map<string, { item: Item; revision: number; bytes?: Uint8Array; offset: number }>();
	let sequence = 0;
	let access: 'read' | 'write' = 'write';
	const shareUrl = 'https://1drv.ms/p/test-fixture';
	const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
	const metadata = (item: Item) => ({ id: item.id, name: item.name, eTag: `"v${item.revision}"`,
		parentReference: { driveId: 'drive', id: 'parent' }, file: {}, webUrl: shareUrl,
		'@microsoft.graph.downloadUrl': `https://transfer.test/download/${item.id}` });
	const requests: Array<{ url: string; init: RequestInit | undefined }> = [];
	const fetch: typeof globalThis.fetch = async (input, init) => {
		init?.signal?.throwIfAborted();
		const url = String(input); requests.push({ url, init });
		// Personal OneDrive can leave the download URL out of a response with $select.
		const shaped = (item: Item) => {
			const value: Record<string, unknown> = metadata(item);
			if (options.selectDropsDownloadUrl && url.includes('$select'))
				delete value['@microsoft.graph.downloadUrl'];
			return value;
		};
		if (url.includes('/shares/')) return json(shaped(first));
		const downloaded = /\/download\/([^/]+)$/u.exec(url)?.[1];
		if (downloaded) return new Response(items.get(downloaded)!.bytes.slice());
		const id = /\/items\/([^/:?]+)(?:\/|\?|$)/u.exec(url)?.[1];
		const found = id ? items.get(decodeURIComponent(id)) : undefined;
		if (url.endsWith('/createLink')) return json({ link: { webUrl: shareUrl } });
		if (url.endsWith('/createUploadSession')) {
			if (access !== 'write') return json({}, 403);
			let item = found;
			if (/\/items\/[^/]+:\//u.test(url)) {
				const name = decodeURIComponent(/:\/([^/]+):\/createUploadSession$/u.exec(url)![1]!);
				item = { id: `copy-${++sequence}`, name, bytes: new Uint8Array(), revision: 1 };
			} else if (new Headers(init?.headers).get('If-Match') !== metadata(item!).eTag) return json({}, 412);
			const uploadUrl = `https://transfer.test/upload/${++sequence}`;
			uploads.set(uploadUrl, { item: item!, revision: item!.revision, offset: 0 });
			return json({ uploadUrl });
		}
		if (url.includes('/upload/')) {
			const upload = uploads.get(url)!;
			if (upload.item.revision !== upload.revision) return json({}, 412);
			const range = /^bytes (\d+)-(\d+)\/(\d+)$/u.exec(new Headers(init?.headers).get('Content-Range') ?? '');
			if (!range || !(init?.body instanceof Uint8Array)) return json({}, 400);
			const [start, end, total] = range.slice(1).map(Number) as [number, number, number];
			if (start !== upload.offset || end - start + 1 !== init.body.length) return json({}, 416);
			upload.bytes ??= new Uint8Array(total);
			upload.bytes.set(init.body, start); upload.offset = end + 1;
			if (upload.offset < total) return json({ nextExpectedRanges: [`${upload.offset}-`] }, 202);
			upload.item.bytes = new Uint8Array(upload.bytes); upload.item.revision++;
			items.set(upload.item.id, upload.item);
			return json(metadata(upload.item));
		}
		if (found) return json(shaped(found));
		return json({}, 404);
	};
	return {
		client: createOneDriveClient({ fetch, getAccessToken: async () => 'test-token', canWrite: () => access === 'write' }),
		shareUrl, requests,
		changeRemotely(bytes = first.bytes) { first.bytes = bytes.slice(); first.revision++; },
		setAccess(value: 'read' | 'write') { access = value; },
		getBytes(id = first.id) { return items.get(id)!.bytes.slice(); },
	};
}
