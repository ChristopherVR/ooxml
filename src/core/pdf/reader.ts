/**
 * Opens a PDF and reads what each page shows (`PdfPageContent`). The file's JavaScript, launch
 * actions and embedded files are never run or opened; only page content, fonts, pictures, link
 * annotations and the document information are read.
 */
import { ContentReader, type Matrix } from './content';
import type { PdfPermissions } from './crypt';
import { PdfDocument, type PdfOpenOptions, type PdfPageNode } from './document';
import { decodeTextString, isName, isNumber, PdfRef, PdfStream, type PdfValue } from './objects';
import type { PdfPageContent, PdfPageLink } from './page-content';

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
	/** What page `index` (0-based) shows. */
	page(index: number, options?: { images?: boolean }): Promise<PdfPageContent>;
}

/** The matrix from user space to reading space: crop box origin, rotation, y down. */
export function pageMatrix(node: Pick<PdfPageNode, 'cropBox' | 'rotate'>): {
	matrix: Matrix;
	width: number;
	height: number;
} {
	const [x0, y0, x1, y1] = node.cropBox as [number, number, number, number];
	const w = x1 - x0;
	const h = y1 - y0;
	switch (node.rotate) {
		case 90:
			return { matrix: [0, 1, 1, 0, -y0, -x0], width: h, height: w };
		case 180:
			return { matrix: [-1, 0, 0, 1, x1, -y0], width: w, height: h };
		case 270:
			return { matrix: [0, -1, -1, 0, y1, x1], width: h, height: w };
		default:
			return { matrix: [1, 0, 0, -1, -x0, y1], width: w, height: h };
	}
}

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
	pageRefs: Map<string, number>,
): Promise<PdfPageLink[]> {
	const annots = await document.resolve(node.dict.get('Annots'));
	if (!Array.isArray(annots)) return [];
	const out: PdfPageLink[] = [];
	for (const item of annots.slice(0, 2000)) {
		const annot = await document.dict(item);
		if (!annot || !isName(annot.get('Subtype'), 'Link')) continue;
		const rect = await document.resolve(annot.get('Rect'));
		if (!Array.isArray(rect) || rect.length !== 4 || !rect.every(isNumber)) continue;
		const [ax, ay, bx, by] = rect as number[];
		const p = [
			[
				matrix[0] * ax! + matrix[2] * ay! + matrix[4],
				matrix[1] * ax! + matrix[3] * ay! + matrix[5],
			],
			[
				matrix[0] * bx! + matrix[2] * by! + matrix[4],
				matrix[1] * bx! + matrix[3] * by! + matrix[5],
			],
		];
		const link: PdfPageLink = {
			x: Math.min(p[0]![0]!, p[1]![0]!),
			y: Math.min(p[0]![1]!, p[1]![1]!),
			width: Math.abs(p[1]![0]! - p[0]![0]!),
			height: Math.abs(p[1]![1]! - p[0]![1]!),
		};
		const action = await document.dict(annot.get('A'));
		if (action && isName(action.get('S'), 'URI')) {
			const uri = decodeTextString(await document.resolve(action.get('URI')));
			if (uri && /^(https?:|mailto:)/i.test(uri)) link.uri = uri;
		}
		const destination = await document.resolve(
			annot.get('Dest') ??
				(action && isName(action.get('S'), 'GoTo') ? action.get('D') : undefined),
		);
		if (Array.isArray(destination) && destination[0] instanceof PdfRef) {
			const page = pageRefs.get(destination[0].key);
			if (page !== undefined) link.page = page;
		}
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
	return {
		pageCount: document.pages.length,
		permissions: document.permissions,
		encrypted: document.decryptor !== undefined,
		signed: isNumber(sigFlags) && (sigFlags & 1) === 1,
		tagged: !!catalog?.has('StructTreeRoot') || markInfo?.get('Marked') === true,
		info,
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
			return {
				index,
				width,
				height,
				spans: reader.result.spans,
				images: reader.result.images,
				rules: reader.result.rules,
				links: await links(document, node, matrix, pageRefs),
				notes: [...reader.result.notes],
			};
		},
	};
}
