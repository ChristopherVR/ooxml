/**
 * A test-only stand-in for Microsoft Graph, answered by Playwright routing. The demos always run
 * the production `ooxml-core/embedding/onedrive` client; only the network is replaced. It answers
 * the drive, folder listings, recent files, item metadata (with `$select`), downloads and upload
 * sessions that File > Open and File > Save make.
 */
import type { BrowserContext, Page } from '@playwright/test';

export const ONEDRIVE_TEST_TOKEN = 'test-only-token';
const OWN_DRIVE = 'own-drive';

interface CloudItem {
	id: string;
	name: string;
	driveId: string;
	/** The folder holding the item; the drive root when absent. */
	parent?: string;
	folder?: boolean;
	bytes: Buffer;
	revision: number;
}

export interface GraphSeed {
	name: string;
	bytes?: Buffer;
	driveId?: string;
	parent?: string;
	folder?: boolean;
}

export interface OneDriveGraphDouble {
	readonly items: Map<string, CloudItem>;
	/** Ids of the items every completed upload replaced, in order. */
	readonly writes: string[];
	/** Make every upload session fail with this HTTP status (0 restores normal uploads). */
	failUploads(status: number): void;
}

export async function routeOneDriveGraph(
	target: Page | BrowserContext,
	seeds: Record<string, GraphSeed>,
): Promise<OneDriveGraphDouble> {
	const items = new Map<string, CloudItem>(
		Object.entries(seeds).map(([id, seed]) => [
			id,
			{
				id,
				name: seed.name,
				driveId: seed.driveId ?? OWN_DRIVE,
				...(seed.parent ? { parent: seed.parent } : {}),
				...(seed.folder ? { folder: true } : {}),
				bytes: seed.bytes ?? Buffer.alloc(0),
				revision: 1,
			},
		]),
	);
	const writes: string[] = [];
	const uploads = new Map<string, { item: CloudItem; chunks: Buffer[] }>();
	let failStatus = 0;
	const describe = (item: CloudItem) => ({
		id: item.id,
		name: item.name,
		eTag: `"revision-${item.revision}"`,
		size: item.bytes.length,
		lastModifiedDateTime: '2026-11-05T09:00:00Z',
		parentReference: { driveId: item.driveId, id: item.parent ?? 'root' },
		...(item.folder
			? { folder: { childCount: [...items.values()].filter((i) => i.parent === item.id).length } }
			: { file: {} }),
		'@microsoft.graph.downloadUrl': `https://transfer.onedrive-test.invalid/download/${item.id}`,
	});
	const children = (parent: string | undefined) =>
		[...items.values()].filter((i) => i.driveId === OWN_DRIVE && i.parent === parent);
	const page = (list: CloudItem[]) => ({ value: list.map(describe) });

	await target.route('https://graph.microsoft.com/v1.0/**', async (route) => {
		const request = route.request();
		if (request.headers().authorization !== `Bearer ${ONEDRIVE_TEST_TOKEN}`) {
			return route.fulfill({ status: 401, json: {} });
		}
		const path = new URL(request.url()).pathname.slice('/v1.0'.length);
		if (path === '/me/drive') return route.fulfill({ json: { id: OWN_DRIVE } });
		if (path === '/me/drive/root/children')
			return route.fulfill({ json: page(children(undefined)) });
		if (path === '/me/drive/recent') {
			// Recent items point at the real item through `remoteItem`, on whichever drive holds it.
			const recent = [...items.values()].filter((i) => !i.folder);
			return route.fulfill({
				json: {
					value: recent.map((item) => ({
						id: `recent-${item.id}`,
						name: item.name,
						remoteItem: describe(item),
					})),
				},
			});
		}
		const match = /^\/drives\/([^/]+)\/items\/([^/]+)(.*)$/u.exec(path);
		const item = match ? items.get(decodeURIComponent(match[2]!)) : undefined;
		if (!match || !item) return route.fulfill({ status: 404, json: {} });
		const rest = match[3];
		if (request.method() === 'GET' && rest === '/children') {
			return route.fulfill({ json: page(children(item.id)) });
		}
		if (request.method() === 'GET' && rest === '') return route.fulfill({ json: describe(item) });
		if (rest === '/createUploadSession') {
			if (failStatus) return route.fulfill({ status: failStatus, json: {} });
			if (request.headers()['if-match'] !== describe(item).eTag) {
				return route.fulfill({ status: 412, json: {} });
			}
			uploads.set(item.id, { item, chunks: [] });
			return route.fulfill({
				json: { uploadUrl: `https://transfer.onedrive-test.invalid/upload/${item.id}` },
			});
		}
		return route.fulfill({ status: 404, json: {} });
	});
	await target.route('https://transfer.onedrive-test.invalid/**', async (route) => {
		const request = route.request();
		const [, action, id] = new URL(request.url()).pathname.split('/');
		if (action === 'download') return route.fulfill({ body: items.get(id!)!.bytes });
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
		upload.item.bytes = Buffer.concat(upload.chunks);
		upload.item.revision += 1;
		writes.push(upload.item.id);
		return route.fulfill({ json: describe(upload.item) });
	});
	return {
		items,
		writes,
		failUploads(status) {
			failStatus = status;
		},
	};
}
