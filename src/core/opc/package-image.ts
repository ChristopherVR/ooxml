/**
 * Adds a picture to a package for one source part: a new media part under the host's folder
 * (`ppt/media/`, `word/media/`, `xl/media/`, `visio/media/`), the source part's `image`
 * relationship and the extension's `Default` content type. A picture whose bytes the source part
 * already references reuses that relationship, so saving the same edit twice writes no duplicate.
 * Every product's chart picture fills allocate through this (`chart/picture-parts.ts`).
 *
 * @module opc/package-image
 */
import type JSZip from 'jszip';
import { relativePartTarget } from './clone-part-graph';
import { ensureRelationship } from './package';
import { RELATIONSHIP_TYPES } from './relationship-types';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from './relationships';

/** Raster and vector picture formats a DrawingML `a:blip` may name, by file extension. */
export const PACKAGE_IMAGE_CONTENT_TYPES: Readonly<Record<string, string>> = {
	png: 'image/png',
	jpg: 'image/jpeg',
	jpeg: 'image/jpeg',
	gif: 'image/gif',
	bmp: 'image/bmp',
	tif: 'image/tiff',
	tiff: 'image/tiff',
	svg: 'image/svg+xml',
	webp: 'image/webp',
	emf: 'image/x-emf',
	wmf: 'image/x-wmf',
};

/** Picture bytes and their file extension (`png`, `jpg`...). */
export interface PackageImage {
	bytes: Uint8Array;
	extension: string;
}

const EXTENSION_OF: Readonly<Record<string, string>> = {
	'image/png': 'png',
	'image/jpeg': 'jpg',
	'image/jpg': 'jpg',
	'image/gif': 'gif',
	'image/bmp': 'bmp',
	'image/tiff': 'tiff',
	'image/svg+xml': 'svg',
	'image/webp': 'webp',
	'image/x-emf': 'emf',
	'image/x-wmf': 'wmf',
};

/** Decodes a base64 `data:image/...` URL; undefined for any other URL or picture type. */
export function decodeImageDataUrl(url: string): PackageImage | undefined {
	const match = /^data:([^;,]+);base64,([A-Za-z0-9+/=\s]*)$/u.exec(url);
	const extension = match && EXTENSION_OF[match[1]!.toLowerCase()];
	if (!match || !extension) return undefined;
	try {
		const binary = atob(match[2]!.replace(/\s/gu, ''));
		const bytes = new Uint8Array(binary.length);
		for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
		return bytes.length ? { bytes, extension } : undefined;
	} catch {
		return undefined;
	}
}

const equalBytes = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((value, index) => value === b[index]);

/** Adds `<Default Extension>` for the picture's extension unless one exists (string surgery). */
async function ensureImageDefault(zip: JSZip, extension: string): Promise<void> {
	const contentType = PACKAGE_IMAGE_CONTENT_TYPES[extension];
	const xml = await zip.file('[Content_Types].xml')?.async('string');
	if (!contentType || !xml) return;
	const declared = [...xml.matchAll(/<Default\b[^>]*\bExtension\s*=\s*["']([^"']*)["']/gu)].some(
		(match) => match[1]!.toLowerCase() === extension,
	);
	if (!declared)
		zip.file(
			'[Content_Types].xml',
			xml.replace(
				/(<Types\b[^>]*>)/u,
				`$1<Default Extension="${extension}" ContentType="${contentType}"/>`,
			),
		);
}

/**
 * Stores `image` for `sourcePart` and returns the relationship id its `r:embed` names. The media
 * part is the first free `<mediaFolder>imageN.<ext>`; an identical picture the part already
 * references keeps its relationship. Throws for an extension without an image content type.
 */
export async function embedPackageImage(
	zip: JSZip,
	sourcePart: string,
	mediaFolder: string,
	image: PackageImage,
): Promise<string> {
	const extension =
		image.extension.toLowerCase() === 'jpeg' ? 'jpg' : image.extension.toLowerCase();
	if (!PACKAGE_IMAGE_CONTENT_TYPES[extension])
		throw new RangeError(`Unsupported picture type: ${image.extension}`);
	const relsPath = relationshipsPartFor(sourcePart);
	const rels = parseRelationships(await zip.file(relsPath)?.async('string'));
	for (const rel of rels.values()) {
		if (rel.type !== RELATIONSHIP_TYPES.image || rel.mode === 'External') continue;
		const existing = await zip.file(resolvePartPath(sourcePart, rel.target))?.async('uint8array');
		if (existing && equalBytes(existing, image.bytes)) return rel.id;
	}
	const folder = mediaFolder.endsWith('/') ? mediaFolder : `${mediaFolder}/`;
	const taken = new Set(Object.keys(zip.files).map((name) => name.toLowerCase()));
	let n = 1;
	while (taken.has(`${folder}image${n}.${extension}`.toLowerCase())) n++;
	const partName = `${folder}image${n}.${extension}`;
	zip.file(partName, image.bytes);
	await ensureImageDefault(zip, extension);
	return ensureRelationship(
		zip,
		relsPath,
		RELATIONSHIP_TYPES.image,
		relativePartTarget(sourcePart, partName),
	);
}
