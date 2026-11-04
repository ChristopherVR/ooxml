/**
 * A test-only stand-in for Microsoft Graph, answered by Playwright routing. The demo always runs
 * the production `ooxml-core/embedding/onedrive` client; only the network is replaced. Mirrors
 * `scripts/test-support/onedrive-http.mjs`, which the Office app's browser check uses.
 */
import type { BrowserContext, Page } from '@playwright/test';

export const ONEDRIVE_TEST_TOKEN = 'test-only-token';
const OWN_DRIVE = 'own-drive';

interface CloudFile {
	id: string;
	name: string;
	driveId: string;
	bytes: Buffer;
	revision: number;
}

export interface OneDriveGraphDouble {
	readonly files: Map<string, CloudFile>;
	/** Ids of the items every completed upload replaced, in order. */
	readonly writes: string[];
	/** Make every upload fail with this HTTP status (0 restores normal uploads). */
	failUploads(status: number): void;
}

/** A sharing link that the double resolves to `id`. */
export const sharingLink = (id: string) => `https://contoso-my.sharepoint.com/:p:/g/personal/${id}`;

export async function routeOneDriveGraph(
	target: Page | BrowserContext,
	decks: Record<string, { name: string; bytes: Buffer; driveId?: string }>,
): Promise<OneDriveGraphDouble> {
	const files = new Map<string, CloudFile>(
		Object.entries(decks).map(([id, deck]) => [
			id,
			{ id, name: deck.name, driveId: deck.driveId ?? OWN_DRIVE, bytes: deck.bytes, revision: 1 },
		]),
	);
	const writes: string[] = [];
	const uploads = new Map<string, { file: CloudFile; chunks: Buffer[] }>();
	let failStatus = 0;
	const metadata = (file: CloudFile) => ({
		id: file.id,
		name: file.name,
		eTag: `"revision-${file.revision}"`,
		parentReference: { driveId: file.driveId, id: 'folder' },
		file: {},
		size: file.bytes.length,
		'@microsoft.graph.downloadUrl': `https://transfer.onedrive-test.invalid/download/${file.id}`,
	});
	await target.route('https://graph.microsoft.com/v1.0/**', async (route) => {
		const request = route.request();
		if (request.headers().authorization !== `Bearer ${ONEDRIVE_TEST_TOKEN}`) {
			return route.fulfill({ status: 401, json: {} });
		}
		const path = new URL(request.url()).pathname.slice('/v1.0'.length);
		if (path === '/me/drive') {
			return route.fulfill({ json: { id: OWN_DRIVE } });
		}
		if (path.startsWith('/shares/')) {
			const token = path.split('/')[2]!;
			const shared = new URL(Buffer.from(token.slice(2), 'base64url').toString());
			const file = files.get(shared.pathname.split('/').pop()!);
			return file
				? route.fulfill({ json: metadata(file) })
				: route.fulfill({ status: 404, json: {} });
		}
		const match = /^\/drives\/([^/]+)\/items\/([^/]+)(.*)$/u.exec(path);
		const file = match ? files.get(decodeURIComponent(match[2]!)) : undefined;
		if (!match || !file) {
			return route.fulfill({ status: 404, json: {} });
		}
		if (request.method() === 'GET') {
			return route.fulfill({ json: metadata(file) });
		}
		if (match[3] === '/createUploadSession') {
			if (failStatus) {
				return route.fulfill({ status: failStatus, json: {} });
			}
			if (request.headers()['if-match'] !== metadata(file).eTag) {
				return route.fulfill({ status: 412, json: {} });
			}
			uploads.set(file.id, { file, chunks: [] });
			return route.fulfill({
				json: { uploadUrl: `https://transfer.onedrive-test.invalid/upload/${file.id}` },
			});
		}
		return route.fulfill({ status: 404, json: {} });
	});
	await target.route('https://transfer.onedrive-test.invalid/**', async (route) => {
		const request = route.request();
		const [, action, id] = new URL(request.url()).pathname.split('/');
		if (action === 'download') {
			return route.fulfill({ body: files.get(id!)!.bytes });
		}
		const upload = uploads.get(id!);
		const range = /^bytes (\d+)-(\d+)\/(\d+)$/u.exec(request.headers()['content-range'] ?? '');
		if (!upload || !range) {
			return route.fulfill({ status: 400, json: {} });
		}
		upload.chunks.push(request.postDataBuffer() ?? Buffer.alloc(0));
		const received = upload.chunks.reduce((total, chunk) => total + chunk.length, 0);
		if (received < Number(range[3])) {
			return route.fulfill({ status: 202, json: { nextExpectedRanges: [`${received}-`] } });
		}
		upload.file.bytes = Buffer.concat(upload.chunks);
		upload.file.revision += 1;
		writes.push(upload.file.id);
		return route.fulfill({ json: metadata(upload.file) });
	});
	return {
		files,
		writes,
		failUploads(status) {
			failStatus = status;
		},
	};
}
