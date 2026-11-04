import { cloudFileKind } from '../cloud-files';
import type { CloudEntry, CloudListing } from '../cloud-files';
import { graphUrl } from './transport';
import type { OneDriveItemRef, OneDriveListOptions } from './types';
import { OneDriveError } from './types';

/** The fields a listing needs; keeps Graph responses small. */
export const LIST_SELECT =
	'id,name,size,folder,file,package,remoteItem,lastModifiedDateTime,lastModifiedBy,parentReference';
/** The fields an open needs, including the short-lived pre-authenticated download URL. */
export const OPEN_SELECT =
	'id,name,size,eTag,webUrl,folder,file,package,remoteItem,parentReference,@microsoft.graph.downloadUrl';

/**
 * Entry ids are `${driveId}!${itemId}` with `!` percent-encoded inside each part (SharePoint drive
 * ids and personal item ids contain `!`), so the first `!` always separates the parts.
 */
const part = (value: string) => encodeURIComponent(value).replace(/!/gu, '%21');
export const encodeEntryId = (ref: OneDriveItemRef) => `${part(ref.driveId)}!${part(ref.itemId)}`;
export function decodeEntryId(id: string): OneDriveItemRef {
	const at = id.indexOf('!');
	try {
		const driveId = decodeURIComponent(id.slice(0, at));
		const itemId = decodeURIComponent(id.slice(at + 1));
		if (at > 0 && driveId && itemId) return { driveId, itemId };
	} catch {
		// Fall through.
	}
	throw new OneDriveError('unsupported', 'This entry does not belong to OneDrive.');
}

export const itemUrl = (ref: OneDriveItemRef) =>
	`/drives/${encodeURIComponent(ref.driveId)}/items/${encodeURIComponent(ref.itemId)}`;

const record = (value: unknown): Record<string, unknown> | undefined =>
	value && typeof value === 'object' ? (value as Record<string, unknown>) : undefined;
const text = (value: unknown) => (typeof value === 'string' && value ? value : undefined);

/** "/drive/root:/Reports/Q1" or "/drives/id/root:/Reports" to "My files > Reports > Q1". */
function location(path: string | undefined): string | undefined {
	if (!path) return undefined;
	const [, rest = ''] = path.split(':');
	let segments: string[];
	try {
		segments = rest.split('/').filter(Boolean).map(decodeURIComponent);
	} catch {
		segments = rest.split('/').filter(Boolean);
	}
	return ['My files', ...segments].join(' > ');
}

/** Maps one Graph driveItem; null when it has no usable identity. */
export function entryFromItem(data: Record<string, unknown>): CloudEntry | null {
	// Recent and shared items point at the real item through `remoteItem`.
	const source = record(data.remoteItem) ?? data;
	const parent = record(source.parentReference) ?? record(data.parentReference);
	const driveId = text(parent?.driveId) ?? text(record(data.parentReference)?.driveId);
	const itemId = text(source.id);
	const name = text(data.name) ?? text(source.name);
	if (!driveId || !itemId || !name) return null;
	const folder = record(source.folder) ?? record(data.folder);
	const isPackage = Boolean(source.package ?? data.package);
	const kind = folder || isPackage ? undefined : cloudFileKind(name);
	const size = typeof source.size === 'number' ? source.size : undefined;
	const modifiedAt = Date.parse(
		text(source.lastModifiedDateTime) ?? text(data.lastModifiedDateTime) ?? '',
	);
	const by = text(record(record(source.lastModifiedBy ?? data.lastModifiedBy)?.user)?.displayName);
	const place = location(text(parent?.path));
	const count = folder && typeof folder.childCount === 'number' ? folder.childCount : undefined;
	return {
		id: encodeEntryId({ driveId, itemId }),
		name,
		type: folder ? 'folder' : 'file',
		...(kind ? { kind } : {}),
		...(size !== undefined && !folder ? { size } : {}),
		...(Number.isFinite(modifiedAt) ? { modified: modifiedAt } : {}),
		...(by ? { modifiedBy: by } : {}),
		...(place ? { location: place } : {}),
		...(count !== undefined ? { childCount: count } : {}),
	};
}

export function listingFromPage(data: Record<string, unknown>): CloudListing {
	const values = Array.isArray(data.value) ? (data.value as unknown[]) : [];
	const entries: CloudEntry[] = [];
	for (const value of values) {
		const entry = record(value) ? entryFromItem(value as Record<string, unknown>) : null;
		if (entry) entries.push(entry);
	}
	const link = data['@odata.nextLink'];
	return link === undefined ? { entries } : { entries, next: graphUrl(link) };
}

/** The first-page path, or the validated paging link when `page` is set. */
export function pagePath(path: string, options: OneDriveListOptions | undefined): string {
	if (options?.page) return graphUrl(options.page);
	const top = Math.min(200, Math.max(1, Math.trunc(options?.top ?? 100)));
	return `${path}${path.includes('?') ? '&' : '?'}$select=${LIST_SELECT}&$top=${top}`;
}

/** Graph search text: single quotes doubled, then URL encoded inside the function call. */
export const searchPath = (query: string) =>
	`/me/drive/root/search(q='${encodeURIComponent(query.replace(/'/gu, "''"))}')`;
