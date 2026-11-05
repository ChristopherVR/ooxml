/**
 * A test-only stand-in for Microsoft Graph, answered by Playwright routing. The demo runs the
 * production `ooxml-core/embedding/onedrive` client; only the network is replaced. Supports the
 * calls File > Open makes: the drive, folder listings, item metadata (with `$select`), recent
 * files, downloads and chunked upload sessions with `If-Match`.
 */
import type { BrowserContext, Page } from '@playwright/test';

export const ONEDRIVE_TEST_TOKEN = 'test-only-token';
const OWN_DRIVE = 'own-drive';

interface CloudFile {
	id: string;
	name: string;
	driveId: string;
	/** The id of the folder holding it; absent for the drive root. */
	folder?: string;
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

export interface DoubleFile {
	name: string;
	bytes: Buffer;
	driveId?: string;
	folder?: string;
}

export async function routeOneDriveGraph(
	target: Page | BrowserContext,
	seed: Record<string, DoubleFile>,
	folders: Record<string, string> = {},
): Promise<OneDriveGraphDouble> {
	const files = new Map<string, CloudFile>(
		Object.entries(seed).map(([id, file]) => [
			id,
			{
				id,
				name: file.name,
				driveId: file.driveId ?? OWN_DRIVE,
				...(file.folder ? { folder: file.folder } : {}),
				bytes: file.bytes,
				revision: 1,
			},
		]),
	);
	const writes: string[] = [];
	const uploads = new Map<string, { file: CloudFile; chunks: Buffer[] }>();
	let failStatus = 0;
	const item = (file: CloudFile) => ({
		id: file.id,
		name: file.name,
		eTag: `"revision-${file.revision}"`,
		parentReference: { driveId: file.driveId, id: file.folder ?? 'root' },
		file: {},
		size: file.bytes.length,
		lastModifiedDateTime: '2026-01-15T10:00:00Z',
		'@microsoft.graph.downloadUrl': `https://transfer.onedrive-test.invalid/download/${file.id}`,
	});
	const folderItem = (id: string) => ({
		id,
		name: folders[id],
		parentReference: { driveId: OWN_DRIVE, id: 'root' },
		folder: { childCount: [...files.values()].filter((file) => file.folder === id).length },
		lastModifiedDateTime: '2026-01-15T10:00:00Z',
	});
	await target.route('https://graph.microsoft.com/v1.0/**', async (route) => {
		const request = route.request();
		if (request.headers().authorization !== `Bearer ${ONEDRIVE_TEST_TOKEN}`) {
			return route.fulfill({ status: 401, json: {} });
		}
		const path = new URL(request.url()).pathname.slice('/v1.0'.length);
		if (path === '/me/drive') return route.fulfill({ json: { id: OWN_DRIVE } });
		if (path === '/me/drive/root/children') {
			const root = [...files.values()].filter((file) => file.driveId === OWN_DRIVE && !file.folder);
			return route.fulfill({
				json: { value: [...Object.keys(folders).map(folderItem), ...root.map(item)] },
			});
		}
		if (path === '/me/drive/recent') {
			return route.fulfill({ json: { value: [...files.values()].map(item) } });
		}
		const match = /^\/drives\/([^/]+)\/items\/([^/]+)(.*)$/u.exec(path);
		const id = match ? decodeURIComponent(match[2]!) : '';
		if (match && id in folders) {
			if (match[3] === '/children') {
				const children = [...files.values()].filter((file) => file.folder === id);
				return route.fulfill({ json: { value: children.map(item) } });
			}
			return route.fulfill({ json: folderItem(id) });
		}
		const file = files.get(id);
		if (!match || !file) return route.fulfill({ status: 404, json: {} });
		if (request.method() === 'GET') return route.fulfill({ json: item(file) });
		if (match[3] === '/createUploadSession') {
			if (failStatus) return route.fulfill({ status: failStatus, json: {} });
			if (request.headers()['if-match'] !== item(file).eTag) {
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
		if (action === 'download') return route.fulfill({ body: files.get(id!)!.bytes });
		const upload = uploads.get(id!);
		const range = /^bytes (\d+)-(\d+)\/(\d+)$/u.exec(request.headers()['content-range'] ?? '');
		if (!upload || !range) return route.fulfill({ status: 400, json: {} });
		upload.chunks.push(request.postDataBuffer() ?? Buffer.alloc(0));
		const received = upload.chunks.reduce((total, chunk) => total + chunk.length, 0);
		if (received < Number(range[3])) {
			return route.fulfill({
				status: 202,
				json: { nextExpectedRanges: [`${received}-${Number(range[3]) - 1}`] },
			});
		}
		upload.file.bytes = Buffer.concat(upload.chunks);
		upload.file.revision += 1;
		writes.push(upload.file.id);
		return route.fulfill({ json: item(upload.file) });
	});
	return {
		files,
		writes,
		failUploads(status) {
			failStatus = status;
		},
	};
}
