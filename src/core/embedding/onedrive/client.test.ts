import { Buffer } from 'node:buffer';
import { describe, expect, it } from 'vitest';
import { createOneDriveClient, encodeOneDriveSharingUrl } from './index';
import { createTestDrive } from './__fixtures__/graph';
import type { OneDriveClientOptions } from './types';

const link = 'https://1drv.ms/p/test';
const bytes = new Uint8Array([80, 75, 3, 4]);
const item = {
	id: 'item',
	name: 'Deck.pptx',
	eTag: '"v1"',
	webUrl: link,
	parentReference: { driveId: 'drive' },
	file: {},
	'@microsoft.graph.downloadUrl': 'https://download.example/file?secret=download',
};
const json = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
function client(fetch: typeof globalThis.fetch, options: Partial<OneDriveClientOptions> = {}) {
	return createOneDriveClient({ fetch, getAccessToken: async () => 'secret-token', ...options });
}

describe('OneDrive host adapter', () => {
	it.each(['javascript:alert(1)', 'https://token:secret@tenant.sharepoint.com/file'])(
		'rejects unsafe metadata web URL %s',
		async (webUrl) => {
			await expect(client(async () => json({ ...item, webUrl })).open(link)).rejects.toMatchObject({
				code: 'invalid-response',
			});
		},
	);
	it('encodes UTF-8 sharing links in unpadded base64url', () => {
		const url = 'https://tenant.sharepoint.com/:p:/réunion?e=abc';
		const encoded = encodeOneDriveSharingUrl(url);
		expect(encoded.startsWith('u!')).toBe(true);
		expect(encoded).not.toMatch(/[=+/]/u);
		const decoded = atob(encoded.slice(2).replace(/-/gu, '+').replace(/_/gu, '/'));
		expect(new TextDecoder().decode(Uint8Array.from(decoded, (c) => c.charCodeAt(0)))).toBe(
			new URL(url).href,
		);
	});

	it.each([
		'https://evil.example/file',
		'http://1drv.ms/p/x',
		'https://1drv.ms.evil.example/p/x',
		'https://user:password@tenant.sharepoint.com/x',
		'https://tenant.sharepoint.com:444/x',
		'file:///private.pptx',
		'not a link',
	])('rejects unsupported link %s before authentication', async (url) => {
		let calls = 0;
		const connection = client(async () => {
			calls++;
			return json(item);
		});
		await expect(async () => connection.open(url)).rejects.toMatchObject({ code: 'invalid-link' });
		expect(calls).toBe(0);
	});

	it('authenticates Graph only, keeps transfers credential-free, and defaults to read-only', async () => {
		const requests: Array<{ url: string; init: RequestInit | undefined }> = [];
		const connection = client(async (url, init) => {
			requests.push({ url: String(url), init });
			return requests.length === 1 ? json(item) : new Response(bytes);
		});
		const opened = await connection.open(link);
		expect(opened.bytes).toEqual(bytes);
		expect(opened.file).toMatchObject({
			driveId: 'drive',
			itemId: 'item',
			eTag: '"v1"',
			access: 'unknown',
		});
		expect(new Headers(requests[0]?.init?.headers).get('Authorization')).toBe(
			'Bearer secret-token',
		);
		expect(requests[0]?.init?.redirect).toBe('error');
		expect(new Headers(requests[1]?.init?.headers).has('Authorization')).toBe(false);
		expect(requests[1]?.init?.credentials).toBe('omit');
		expect(requests[1]?.init?.referrerPolicy).toBe('no-referrer');
		await expect(connection.save(opened.file, bytes)).rejects.toMatchObject({
			code: 'access-denied',
		});
	});

	it('opens with one unselected request, and asks again when the download URL is missing', async () => {
		const plain = createTestDrive({ bytes: new Uint8Array([1, 2, 3]) });
		await plain.client.open(plain.shareUrl);
		const plainCalls = plain.requests.filter((request) =>
			request.url.includes('graph.microsoft.com'),
		);
		expect(plainCalls).toHaveLength(1);
		expect(plainCalls[0]!.url).not.toContain('$select');
		const drive = createTestDrive({ bytes: new Uint8Array([1, 2, 3]), omitDownloadUrl: 1 });
		const opened = await drive.client.open(drive.shareUrl);
		expect([...opened.bytes]).toEqual([1, 2, 3]);
		const graphCalls = drive.requests.filter((request) =>
			request.url.includes('graph.microsoft.com'),
		);
		expect(graphCalls).toHaveLength(2);
		expect(graphCalls[1]!.url).not.toContain('$select');
		const transfer = drive.requests.find((request) => request.url.includes('/download/'));
		expect(new Headers(transfer?.init?.headers).get('Authorization')).toBeNull();
	});

	it('does not treat an edit sharing URL as proof of write access', async () => {
		const connection = client(async (url) =>
			String(url).includes('/shares/')
				? json({ ...item, shared: { scope: 'anonymous' } })
				: new Response(bytes),
		);
		const opened = await connection.open(link + '?edit=1');
		expect(opened.file.access).toBe('unknown');
	});

	it('opens and saves through the same transport without a registration, updating the eTag', async () => {
		const mock = createTestDrive({ bytes });
		const opened = await mock.client.open(mock.shareUrl);
		expect(opened.file.access).toBe('write');
		const changed = new Uint8Array([1, 2, 3, 4, 5]);
		const saved = await mock.client.save(opened.file, changed);
		expect(saved.eTag).not.toBe(opened.file.eTag);
		expect(mock.getBytes()).toEqual(changed);
		const reopened = await mock.client.open(mock.shareUrl);
		expect(reopened.bytes).toEqual(changed);
		expect(reopened.file.eTag).toBe(saved.eTag);
	});

	it('rejects stale versions and retains remote bytes', async () => {
		const mock = createTestDrive({ bytes });
		const opened = await mock.client.open(mock.shareUrl);
		mock.changeRemotely(new Uint8Array([9]));
		await expect(mock.client.save(opened.file, bytes)).rejects.toMatchObject({
			code: 'conflict',
			status: 412,
		});
		expect(mock.getBytes()).toEqual(new Uint8Array([9]));
	});

	it('rechecks host access before starting any save', async () => {
		const mock = createTestDrive({ bytes });
		const opened = await mock.client.open(mock.shareUrl);
		mock.setAccess('read');
		await expect(mock.client.save(opened.file, bytes)).rejects.toMatchObject({
			code: 'access-denied',
		});
		expect((await mock.client.open(mock.shareUrl)).file.access).toBe('read');
	});

	it('uses If-Match on session creation and never sends Graph auth to upload URLs', async () => {
		const requests: RequestInit[] = [];
		const connection = client(
			async (url, init = {}) => {
				requests.push(init);
				if (String(url).includes('/shares/')) return json(item);
				if (String(url).includes('/createUploadSession'))
					return json({ uploadUrl: 'https://upload.example/file' });
				if (init.method === 'PUT') return json({ ...item, eTag: '"v2"' });
				return new Response(bytes);
			},
			{ canWrite: () => true },
		);
		const opened = await connection.open(link);
		await connection.save(opened.file, bytes);
		expect(new Headers(requests[2]?.headers).get('If-Match')).toBe('"v1"');
		expect(new Headers(requests[3]?.headers).has('Authorization')).toBe(false);
		expect(new Headers(requests[3]?.headers).get('Content-Range')).toBe('bytes 0-3/4');
	});

	it('uploads sequential 320 KiB aligned fragments', async () => {
		const large = new Uint8Array(5 * 1024 * 1024 + 7).fill(7);
		const mock = createTestDrive({ bytes });
		const opened = await mock.client.open(mock.shareUrl);
		await mock.client.save(opened.file, large);
		// Native byte comparison avoids millions of deep-equality checks on busy CI workers.
		expect(Buffer.from(mock.getBytes()).equals(Buffer.from(large))).toBe(true);
		const fragments = mock.requests.filter(({ url }) => url.includes('/upload/'));
		expect(fragments.length).toBeGreaterThan(1);
		for (const fragment of fragments.slice(0, -1)) {
			const body = fragment.init?.body;
			if (!(body instanceof Uint8Array)) throw new Error('Expected an upload fragment');
			expect(body.length % (320 * 1024)).toBe(0);
		}
	});

	it('bounds streamed downloads even without Content-Length', async () => {
		const connection = client(
			async (url) =>
				String(url).includes('/shares/')
					? json(item)
					: new Response(
							new ReadableStream({
								start(c) {
									c.enqueue(bytes);
									c.close();
								},
							}),
						),
			{ maxBytes: 3 },
		);
		await expect(connection.open(link)).rejects.toMatchObject({ code: 'too-large' });
	});

	it.each([401, 403, 412, 423, 429, 500])('sanitizes HTTP %s errors', async (status) => {
		const connection = client(async () => json({ error: { message: 'secret-token' } }, status));
		await expect(connection.open(link)).rejects.toMatchObject({ status });
		await expect(connection.open(link)).rejects.not.toThrow('secret-token');
	});

	it('sanitizes token-provider and fetch error messages', async () => {
		await expect(
			client(async () => {
				throw new Error('secret-token');
			}).open(link),
		).rejects.toThrow('Could not connect');
		await expect(
			client(async () => json(item), {
				getAccessToken: async () => {
					throw new Error('secret-token');
				},
			}).open(link),
		).rejects.toMatchObject({ code: 'authentication' });
	});

	it('supports cancellation before any request and during token acquisition', async () => {
		const abort = new AbortController();
		abort.abort();
		await expect(
			client(async () => json(item)).open(link, { signal: abort.signal }),
		).rejects.toMatchObject({ name: 'AbortError' });
		const later = new AbortController();
		const pending = client(async () => json(item), {
			getAccessToken: () => new Promise(() => {}),
		}).open(link, { signal: later.signal });
		later.abort();
		await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
	});

	it('rejects unsafe transfer URLs and incomplete identities', async () => {
		await expect(
			client(async () =>
				json({ ...item, '@microsoft.graph.downloadUrl': 'http://unsafe.example/file' }),
			).open(link),
		).rejects.toMatchObject({ code: 'invalid-response' });
		await expect(
			client(async () => json({ ...item, eTag: undefined })).open(link),
		).rejects.toMatchObject({ code: 'invalid-response' });
		await expect(
			client(async () => json({ ...item, name: 'not-an-office-file.pdf' })).open(link),
		).rejects.toMatchObject({ code: 'unsupported' });
	});

	it('snapshots bytes and identity before an asynchronous permission check', async () => {
		let release: (() => void) | undefined;
		let delay = false;
		let postedPath = '';
		let uploaded = new Uint8Array();
		const connection = client(
			async (url, init) => {
				if (String(url).includes('/shares/')) return json(item);
				if (String(url).includes('/createUploadSession')) {
					postedPath = String(url);
					return json({ uploadUrl: 'https://upload.example/file' });
				}
				if (init?.method === 'PUT') {
					uploaded = new Uint8Array(init.body as Uint8Array);
					return json({ ...item, eTag: '"v2"' });
				}
				return new Response(bytes);
			},
			{
				canWrite: async () => {
					if (delay)
						await new Promise<void>((resolve) => {
							release = resolve;
						});
					return true;
				},
			},
		);
		const opened = await connection.open(link);
		delay = true;
		const local = bytes.slice();
		const pending = connection.save(opened.file, local);
		local.fill(0);
		opened.file.itemId = 'changed';
		delay = false;
		release?.();
		await pending;
		expect(postedPath).toContain('/items/item/createUploadSession');
		expect(uploaded).toEqual(bytes);
	});

	it('validates save file metadata before making an upload request', async () => {
		let count = 0;
		const connection = client(
			async (url) => {
				count++;
				return String(url).includes('/shares/') ? json(item) : new Response(bytes);
			},
			{ canWrite: () => true },
		);
		const opened = await connection.open(link);
		await expect(connection.save({ ...opened.file, name: 'bad.pdf' }, bytes)).rejects.toMatchObject(
			{ code: 'unsupported' },
		);
		await expect(connection.save({ ...opened.file, driveId: '' }, bytes)).rejects.toMatchObject({
			code: 'invalid-response',
		});
		expect(count).toBe(2);
	});
});
