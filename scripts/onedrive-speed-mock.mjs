/**
 * A local in-process stand-in for Microsoft Graph and its download host, for
 * onedrive-speed-check.mjs --mock. It contains no Microsoft data and needs no network.
 */
import { createServer } from 'node:http';

const GRAPH_HOST = 'graph.microsoft.com';
const DOWNLOAD_HOST = 'download.mock.invalid';
const DRIVE = 'b!mockdrive';
const ITEM = 'MOCKITEM1';
const FOLDER = 'MOCKFOLDER1';
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * @param {{ size: number, latency: number, bandwidth: number }} options size in MB, latency in ms,
 *   bandwidth in MB/s for the download host (0 = unlimited).
 */
export async function startMockGraph({ size, latency, bandwidth }) {
	const total = Math.max(1, Math.round(size * 1048576));
	const payload = Buffer.alloc(total, 0x50);
	let nonce = 0;
	let origin = '';
	const file = (name) => ({
		id: ITEM,
		name,
		size: total,
		eTag: '"mock-etag,1"',
		webUrl: 'https://contoso.sharepoint.com/mock/' + encodeURIComponent(name),
		file: { mimeType: 'application/octet-stream' },
		parentReference: { driveId: DRIVE, id: FOLDER },
		// Fresh per request, like the real short-lived pre-authenticated URL.
		'@microsoft.graph.downloadUrl': `https://${DOWNLOAD_HOST}/dl/${++nonce}`,
	});
	const entry = (id, name) => ({
		id,
		name,
		size: 4096,
		file: { mimeType: 'application/octet-stream' },
		lastModifiedDateTime: '2026-01-01T00:00:00Z',
		parentReference: { driveId: DRIVE, id: FOLDER },
	});
	const listing = () => ({
		value: Array.from({ length: 25 }, (_, i) => entry(`MOCKCHILD${i}`, `file-${i}.docx`)),
	});
	const json = (res, body, status = 200) => {
		const text = JSON.stringify(body);
		res.writeHead(status, {
			'Content-Type': 'application/json',
			'Content-Length': Buffer.byteLength(text),
		});
		res.end(text);
	};
	async function download(res) {
		res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Length': total });
		const chunk = 64 * 1024;
		const started = performance.now();
		for (let at = 0; at < total; at += chunk) {
			if (!res.write(payload.subarray(at, Math.min(total, at + chunk))))
				await new Promise((resolve) => res.once('drain', resolve));
			if (bandwidth > 0) {
				const ahead =
					(Math.min(total, at + chunk) / (bandwidth * 1048576)) * 1000 -
					(performance.now() - started);
				if (ahead > 1) await wait(ahead);
			}
		}
		res.end();
	}
	const server = createServer(async (req, res) => {
		try {
			if (latency > 0) await wait(latency);
			const url = new URL(req.url ?? '/', 'http://localhost');
			const path = decodeURIComponent(url.pathname);
			if (path.startsWith('/dl/')) return await download(res);
			const content = path.endsWith('/content');
			const base = content ? path.slice(0, -'/content'.length).replace(/:$/u, '') : path;
			if (
				content &&
				/^\/v1\.0\/(shares\/.+\/driveItem|me\/drive\/root:\/.+|drives\/.+\/items\/.+)$/u.test(base)
			) {
				// Real Graph answers /content with a 302 to the download host.
				res.writeHead(302, {
					Location: `https://${DOWNLOAD_HOST}/dl/${++nonce}`,
					'Content-Length': 0,
				});
				return res.end();
			}
			if (/\/children$/u.test(path) || path.endsWith('/me/drive/recent'))
				return json(res, listing());
			if (/^\/v1\.0\/shares\/[^/]+\/driveItem$/u.test(path)) return json(res, file('speed.pptx'));
			if (/^\/v1\.0\/me\/drive\/root:\/.+/u.test(path)) {
				const name =
					path
						.slice(path.indexOf(':/') + 2)
						.split('/')
						.pop() ?? '';
				return /\.(docx|xlsx|pptx|vsdx)$/iu.test(name)
					? json(res, file(name))
					: json(res, {
							id: FOLDER,
							name,
							folder: { childCount: 25 },
							parentReference: { driveId: DRIVE },
						});
			}
			if (/^\/v1\.0\/drives\/[^/]+\/items\/[^/]+$/u.test(path))
				return json(res, file('speed.pptx'));
			json(res, { error: { code: 'itemNotFound' } }, 404);
		} catch {
			res.destroy();
		}
	});
	await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
	origin = `http://127.0.0.1:${server.address().port}`;
	/** A fetch that sends the library's Graph and download-host URLs to the local server. */
	const mockFetch = (input, init) => {
		const url = new URL(String(input));
		if (url.hostname !== GRAPH_HOST && url.hostname !== DOWNLOAD_HOST)
			throw new Error('The mock only serves its own hosts.');
		return fetch(origin + url.pathname + url.search, init);
	};
	return {
		fetch: mockFetch,
		shareUrl: 'https://1drv.ms/p/s!mockshare',
		path: 'Documents/speed.pptx',
		folder: 'Documents',
		close: () =>
			new Promise((resolve) => {
				server.closeAllConnections?.();
				server.close(resolve);
			}),
	};
}
