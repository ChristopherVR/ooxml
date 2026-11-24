/**
 * Opens a PDF and reads what each page shows (`PdfPageContent`). The file's JavaScript, launch
 * actions and embedded files are never run or opened; only page content, fonts, pictures, link
 * annotations, destinations, the outline, the structure tree and the document information are
 * read.
 */
import { ContentReader, type Matrix } from './content';
import type { PdfPermissions } from './crypt-types';
import { PdfDocument, type PdfOpenOptions, type PdfPageNode } from './document';
import { pageMatrix } from './matrix';
import { Destinations } from './navigation';
import { decodeTextString, isName, isNumber, PdfStream, type PdfValue } from './objects';
import type { PdfOutlineItem, PdfPageContent, PdfPageLink } from './page-content';
import { readStructure, type StructureMarks } from './structure';

export { pageMatrix };

export interface PdfDocumentInfo {
	title?: string;
	author?: string;
	subject?: string;
	keywords?: string;
	creator?: string;
	producer?: string;
	/** The catalog's `/Lang`, e.g. `fr-FR`. */
	language?: string;
}

export interface PdfReader {
	readonly pageCount: number;
	/** Set when the file is encrypted; what its permission flags allow. */
	readonly permissions: PdfPermissions | undefined;
	readonly encrypted: boolean;
	/** The file holds a digital signature (which a conversion cannot keep). */
	readonly signed: boolean;
	/** The file is tagged (has a structure tree). */
	readonly tagged: boolean;
	readonly info: PdfDocumentInfo;
	/** The document outline (bookmarks), flattened in document order. */
	readonly outline: PdfOutlineItem[];
	/** What page `index` (0-based) shows. */
	page(index: number, options?: { images?: boolean }): Promise<PdfPageContent>;
}

/** Why the reader could not return a painted shape as a path. */
const DRAWING_REASONS = {
	shading: 'shadings (gradients) are not supported',
	pattern: 'pattern fills are not supported',
	clip: 'a clipping path cuts them',
	size: 'they have too many segments',
	count: 'the page has too many shapes',
} as const;

async function contentBytes(
	document: PdfDocument,
	value: PdfValue | undefined,
): Promise<Uint8Array> {
	const resolved = await document.resolve(value);
	const streams = Array.isArray(resolved) ? resolved : [resolved];
	const parts: Uint8Array[] = [];
	for (const item of streams) {
		const stream = await document.resolve(item);
		if (!(stream instanceof PdfStream)) continue;
		parts.push(await document.streamBytes(stream), Uint8Array.of(10));
	}
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
}

async function links(
	document: PdfDocument,
	node: PdfPageNode,
	matrix: Matrix,
	destinations: Destinations,
): Promise<PdfPageLink[]> {
	const annots = await document.resolve(node.dict.get('Annots'));
	if (!Array.isArray(annots)) return [];
	const out: PdfPageLink[] = [];
	for (const item of annots.slice(0, 2000)) {
		const annot = await document.dict(item);
		if (!annot || !isName(annot.get('Subtype'), 'Link')) continue;
		const rect = await document.resolve(annot.get('Rect'));
		if (!Array.isArray(rect) || rect.length !== 4 || !rect.every(isNumber)) continue;
		const [ax, ay, bx, by] = rect as [number, number, number, number];
		const a = [
			matrix[0] * ax + matrix[2] * ay + matrix[4],
			matrix[1] * ax + matrix[3] * ay + matrix[5],
		];
		const b = [
			matrix[0] * bx + matrix[2] * by + matrix[4],
			matrix[1] * bx + matrix[3] * by + matrix[5],
		];
		const link: PdfPageLink = {
			x: Math.min(a[0]!, b[0]!),
			y: Math.min(a[1]!, b[1]!),
			width: Math.abs(b[0]! - a[0]!),
			height: Math.abs(b[1]! - a[1]!),
		};
		const action = await document.dict(annot.get('A'));
		if (action && isName(action.get('S'), 'URI')) {
			const uri = decodeTextString(await document.resolve(action.get('URI')));
			if (uri && /^(https?:|mailto:)/i.test(uri)) link.uri = uri;
		}
		const target = await destinations.linkTarget(annot, action);
		if (target?.page !== undefined) {
			link.page = target.page;
			if (target.y !== undefined) link.targetY = target.y;
		}
		if (target?.destination !== undefined) link.destination = target.destination;
		if (link.uri || link.page !== undefined) out.push(link);
	}
	return out;
}

/** Opens `bytes` as a PDF (throws `PdfPasswordError` for a password it cannot open with). */
export async function openPdf(bytes: Uint8Array, options: PdfOpenOptions = {}): Promise<PdfReader> {
	const document = await PdfDocument.open(bytes, options);
	const catalog = await document.catalog();
	const language = decodeTextString(await document.resolve(catalog?.get('Lang')));
	const raw = await document.info();
	const info: PdfDocumentInfo = {
		...(raw.Title ? { title: raw.Title } : {}),
		...(raw.Author ? { author: raw.Author } : {}),
		...(raw.Subject ? { subject: raw.Subject } : {}),
		...(raw.Keywords ? { keywords: raw.Keywords } : {}),
		...(raw.Creator ? { creator: raw.Creator } : {}),
		...(raw.Producer ? { producer: raw.Producer } : {}),
		...(language?.trim() ? { language: language.trim() } : {}),
	};
	const acroForm = await document.dict(catalog?.get('AcroForm'));
	const sigFlags = await document.resolve(acroForm?.get('SigFlags'));
	const markInfo = await document.dict(catalog?.get('MarkInfo'));
	const pageRefs = new Map<string, number>();
	document.pages.forEach((node, index) => node.ref && pageRefs.set(node.ref.key, index));
	const destinations = new Destinations(document, catalog, pageRefs);
	let outline: PdfOutlineItem[] = [];
	try {
		outline = await destinations.outline();
	} catch {
		// A damaged outline only loses the bookmarks.
	}
	let structure: Promise<StructureMarks> | undefined;
	const marksOf = (index: number) => {
		structure ??= readStructure(document, catalog, pageRefs).catch(() => new Map());
		return structure.then((marks) => marks.get(index));
	};
	return {
		pageCount: document.pages.length,
		permissions: document.permissions,
		encrypted: document.decryptor !== undefined,
		signed: isNumber(sigFlags) && (sigFlags & 1) === 1,
		tagged: !!catalog?.has('StructTreeRoot') || markInfo?.get('Marked') === true,
		info,
		outline,
		async page(index, pageOptions = {}) {
			const node = document.pages[index];
			if (!node) throw new RangeError(`The PDF has no page ${index + 1}.`);
			const { matrix, width, height } = pageMatrix(node);
			const reader = new ContentReader(document, matrix, pageOptions.images !== false);
			try {
				await reader.run(
					await contentBytes(document, node.dict.get('Contents')),
					node.resources,
					[1, 0, 0, 1, 0, 0],
				);
			} catch (error) {
				reader.result.notes.add(
					`The page content could not be read completely (${error instanceof Error ? error.message : 'error'}).`,
				);
			}
			const { result } = reader;
			if (result.marks.size) {
				const marks = await marksOf(index);
				for (const [item, mcid] of result.marks) {
					const mark = marks?.get(mcid);
					if (!mark) continue;
					if ('text' in item) {
						item.role = mark.role;
						item.element = mark.element;
					} else if (mark.alt && !item.alt) item.alt = mark.alt;
				}
			}
			const leftOut = [...result.drawings.values()].reduce((sum, count) => sum + count, 0);
			if (leftOut) {
				const why = [...result.drawings.keys()].map((reason) => DRAWING_REASONS[reason]).join(', ');
				result.notes.add(
					`Vector drawings (${leftOut} shape${leftOut === 1 ? '' : 's'}) were left out: ${why}.`,
				);
			}
			return {
				index,
				width,
				height,
				spans: result.spans,
				images: result.images,
				rules: result.rules,
				paths: result.paths,
				links: await links(document, node, matrix, destinations),
				notes: [...result.notes],
			};
		},
	};
}
