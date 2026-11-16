/**
 * Relationship allocation for chart picture fills (`a:blipFill` on a chart element), shared by
 * every host that saves a chart part. A Format pane hands a picture over as a `data:` URL; the host
 * calls {@link embedChartPictures} with its chart part and media folder before it writes the
 * fills, and puts each returned id in `a:blip/@r:embed` (`chartFillFromAppearance` with
 * `picture.relId`). The same URL twice, or a picture the chart already references, shares one
 * media part and relationship.
 *
 * A `ChartSpace` carries its pictures as URLs in `pictureSources`, keyed by relationship id, so the
 * shared painter draws them: {@link chartSpacePictureSources} reads the pictures a loaded part
 * references, a Format pane edit (`setChartTargetPicture`) files a new one under a pending id, and
 * {@link embedChartSpacePictures} stores the pending ones on save and swaps in their real ids.
 *
 * @module chart/picture-parts
 */
import type JSZip from 'jszip';
import { encodeBase64 } from '../digest/base64';
import {
	decodeImageDataUrl,
	embedPackageImage,
	PACKAGE_IMAGE_CONTENT_TYPES,
	type PackageImage,
} from '../opc/package-image';
import type { Relationship } from '../opc/relationships';
import { resolvePartPath } from '../opc/relationships';
import { RELATIONSHIP_TYPES } from '../opc/relationship-types';
import type { DrawingFill } from '../drawingml/types';
import { isPendingChartPicture } from './edit/picture';
import type { ChartSpace } from './model';

/** The media folder of each host's package, where chart pictures are stored. */
export const CHART_MEDIA_FOLDERS = {
	pptx: 'ppt/media/',
	docx: 'word/media/',
	xlsx: 'xl/media/',
	visio: 'visio/media/',
} as const;

/**
 * Stores each picture for `chartPartName` and returns the relationship id of each source. A
 * source is a `data:image/...` URL or decoded bytes keyed by the caller; a URL that does not
 * decode is left out of the result, so the caller keeps that fill unchanged.
 */
export async function embedChartPictures(
	zip: JSZip,
	chartPartName: string,
	mediaFolder: string,
	sources: Iterable<string>,
	decode: (source: string) => PackageImage | undefined | null = decodeImageDataUrl,
): Promise<Map<string, string>> {
	const ids = new Map<string, string>();
	for (const source of sources) {
		if (ids.has(source)) continue;
		const image = decode(source);
		if (image) ids.set(source, await embedPackageImage(zip, chartPartName, mediaFolder, image));
	}
	return ids;
}

type PictureFill = Extract<DrawingFill, { kind: 'picture' }>;

/** Calls `visit` for every picture fill in `space` (any depth); returns a copy when it maps one. */
function mapPictureFills(
	value: unknown,
	visit: (fill: PictureFill) => PictureFill | undefined,
): unknown {
	if (Array.isArray(value)) {
		let changed = false;
		const out = value.map((item) => {
			const next = mapPictureFills(item, visit);
			changed ||= next !== item;
			return next;
		});
		return changed ? out : value;
	}
	if (!value || typeof value !== 'object') return value;
	const record = value as Record<string, unknown>;
	if (record.kind === 'picture' && typeof record.relId === 'string')
		return visit(record as unknown as PictureFill) ?? value;
	let out: Record<string, unknown> | undefined;
	for (const [key, item] of Object.entries(record)) {
		if (key === 'pictureSources') continue;
		const next = mapPictureFills(item, visit);
		if (next !== item) (out ??= { ...record })[key] = next;
	}
	return out ?? value;
}

/** Every relationship id the picture fills of `space` name. */
export function chartSpacePictureIds(space: ChartSpace): string[] {
	const ids = new Set<string>();
	mapPictureFills(space, (fill) => {
		if (fill.relId) ids.add(fill.relId);
		return undefined;
	});
	return [...ids];
}

/** A picture's bytes as a `data:` URL, by its part's extension; undefined for other parts. */
export function chartPictureDataUrl(partName: string, bytes: Uint8Array): string | undefined {
	const extension = partName.split('.').pop()?.toLowerCase() ?? '';
	const type = PACKAGE_IMAGE_CONTENT_TYPES[extension === 'jpeg' ? 'jpg' : extension];
	return type && bytes.length ? `data:${type};base64,${encodeBase64(bytes)}` : undefined;
}

/**
 * The pictures the fills of a loaded chart part reference, as `data:` URLs by relationship id, for
 * `ChartSpace.pictureSources`. `rels` are the chart part's relationships; `read` returns a part's
 * bytes. Missing, external and unreadable pictures are left out (the painter skips them).
 */
export async function chartSpacePictureSources(
	space: ChartSpace,
	chartPartName: string,
	rels: ReadonlyMap<string, Relationship>,
	read: (partName: string) => Promise<Uint8Array | undefined> | Uint8Array | undefined,
): Promise<Record<string, string>> {
	const sources: Record<string, string> = {};
	for (const id of chartSpacePictureIds(space)) {
		const rel = rels.get(id);
		if (!rel || rel.mode === 'External' || rel.type !== RELATIONSHIP_TYPES.image) continue;
		const part = resolvePartPath(chartPartName, rel.target);
		const bytes = await read(part);
		const url = bytes && chartPictureDataUrl(part, bytes);
		if (url) sources[id] = url;
	}
	return sources;
}

/**
 * Stores the pending pictures of `space` (filed by a Format pane edit) as media parts of
 * `chartPartName` and returns the model to write: the fills name the new relationship ids and
 * `pictureSources` is dropped. Throws a `RangeError` for a pending id with no picture that
 * decodes, so a save never writes a relationship id that resolves to nothing.
 */
export async function embedChartSpacePictures(
	zip: JSZip,
	chartPartName: string,
	mediaFolder: string,
	space: ChartSpace,
): Promise<ChartSpace> {
	const { pictureSources, ...rest } = space;
	const pending = chartSpacePictureIds(rest).filter(isPendingChartPicture);
	if (!pending.length) return pictureSources ? (rest as ChartSpace) : space;
	const urls = pending.map((id) => pictureSources?.[id] ?? '');
	const ids = await embedChartPictures(zip, chartPartName, mediaFolder, urls);
	const real = new Map(pending.map((id, i) => [id, ids.get(urls[i]!)]));
	const missing = pending.find((id) => !real.get(id));
	if (missing) throw new RangeError(`The chart picture ${missing} has no picture to save`);
	return mapPictureFills(rest, (fill) =>
		isPendingChartPicture(fill.relId) ? { ...fill, relId: real.get(fill.relId!)! } : undefined,
	) as ChartSpace;
}

/**
 * Stores the pending pictures a written chart part names (`r:embed="chartPicturePendingN"`) as
 * media parts of that part and rewrites their ids in place, for a host that writes the part
 * before it knows where the pictures go (a chart inserted this session, a package built by its own
 * writer). `sources` is the model's `pictureSources`. Returns whether the part changed; throws a
 * `RangeError` for a pending id with no picture that decodes.
 */
export async function embedChartPartPictures(
	zip: JSZip,
	chartPartName: string,
	mediaFolder: string,
	sources: Readonly<Record<string, string>> | undefined,
): Promise<boolean> {
	const xml = await zip.file(chartPartName)?.async('string');
	if (!xml) return false;
	const pending = [
		...new Set(
			[...xml.matchAll(/\br:(?:embed|link)\s*=\s*"([^"]*)"/gu)]
				.map((match) => match[1]!)
				.filter(isPendingChartPicture),
		),
	];
	if (!pending.length) return false;
	const urls = pending.map((id) => sources?.[id] ?? '');
	const ids = await embedChartPictures(zip, chartPartName, mediaFolder, urls);
	let next = xml;
	pending.forEach((id, i) => {
		const real = ids.get(urls[i]!);
		if (!real) throw new RangeError(`The chart picture ${id} has no picture to save`);
		next = next.replaceAll(`"${id}"`, `"${real}"`);
	});
	zip.file(chartPartName, next);
	return true;
}
