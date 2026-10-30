import type { VisioDocument } from 'ooxml-core/visio';
import { PdfFontSet, type PdfFontFile } from '../export/pdf-embedded-font';
import { buildVectorPdf, PdfVectorPage, type PdfImageResource } from '../export/pdf-vector';
import { exportPageSvg } from './export-svg';
import { collectSvgImages, convertSvgPage } from './pdf-svg';
import { rasterizePage } from './viewer-export';

/** Which pages the PDF holds. Background pages are always drawn behind the pages that use them. */
export interface PdfExportOptions {
	/** `current` exports only `pageIndex` (a foreground page); default `all`. */
	scope?: 'all' | 'current';
	pageIndex?: number;
	/** 1-based inclusive range over the foreground pages. */
	from?: number;
	to?: number;
	/**
	 * TrueType fonts to embed (subset) so text outside WinAnsi, or in a typeface the standard
	 * fonts do not match, stays real text. The host supplies the bytes; none are bundled.
	 */
	fonts?: readonly PdfFontFile[];
}

export interface PdfPageReport {
	/** Index in the document's page list. */
	index: number;
	name: string;
	/** `vector` is real paths and text; `picture` is a raster of the page, with `fallback` saying why. */
	mode: 'vector' | 'picture';
	fallback: string[];
	/** Effects the vector page leaves out or simplifies. */
	degraded: string[];
}

export interface PdfExportResult {
	blob: Blob;
	pages: PdfPageReport[];
	/** Families of the supplied fonts that were embedded as subsets. */
	embeddedFonts: string[];
	/** Supplied fonts that were not embedded (not TrueType, or the licence forbids embedding). */
	refusedFonts: string[];
}

const MAX_IMAGE_SIDE = 2048;
const MAX_PAGES = 2000;
const MAX_IMAGES_PER_PAGE = 500;

async function deflate(bytes: Uint8Array, view: Window): Promise<Uint8Array | undefined> {
	const Compression = (view as unknown as { CompressionStream?: typeof CompressionStream })
		.CompressionStream;
	if (!Compression) return undefined;
	const stream = new Blob([Uint8Array.from(bytes)])
		.stream()
		.pipeThrough(new Compression('deflate'));
	return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Decode a data URL picture to RGB (and alpha) samples, downscaled to a bounded size. */
async function decodeImage(doc: Document, href: string): Promise<PdfImageResource | undefined> {
	const view = doc.defaultView;
	if (!view || !href.startsWith('data:image/')) return undefined;
	try {
		const image = new view.Image();
		await new Promise<void>((resolve, reject) => {
			image.onload = () => resolve();
			image.onerror = () => reject(new Error('picture'));
			image.src = href;
		});
		const scale = Math.min(
			1,
			MAX_IMAGE_SIDE / Math.max(image.naturalWidth, image.naturalHeight, 1),
		);
		const width = Math.max(1, Math.round(image.naturalWidth * scale));
		const height = Math.max(1, Math.round(image.naturalHeight * scale));
		const canvas = doc.createElement('canvas');
		canvas.width = width;
		canvas.height = height;
		const context = canvas.getContext('2d');
		if (!context) return undefined;
		context.drawImage(image, 0, 0, width, height);
		const rgba = context.getImageData(0, 0, width, height).data;
		const rgb = new Uint8Array(width * height * 3);
		const alpha = new Uint8Array(width * height);
		let translucent = false;
		for (let pixel = 0; pixel < width * height; pixel++) {
			rgb[pixel * 3] = rgba[pixel * 4]!;
			rgb[pixel * 3 + 1] = rgba[pixel * 4 + 1]!;
			rgb[pixel * 3 + 2] = rgba[pixel * 4 + 2]!;
			alpha[pixel] = rgba[pixel * 4 + 3]!;
			if (alpha[pixel] !== 255) translucent = true;
		}
		const packed = await deflate(rgb, view);
		const packedAlpha = translucent ? await deflate(alpha, view) : undefined;
		return {
			width,
			height,
			data: packed ?? rgb,
			...(packed ? { filter: 'FlateDecode' as const } : {}),
			...(translucent
				? {
						alpha: {
							data: packedAlpha ?? alpha,
							...(packedAlpha ? { filter: 'FlateDecode' as const } : {}),
						},
					}
				: {}),
		};
	} catch {
		return undefined;
	}
}

/** The foreground page indexes a PDF export covers. */
export function pdfPageIndexes(model: VisioDocument, options: PdfExportOptions = {}): number[] {
	const foreground = model.pages.flatMap((page, index) => (page.isBackground ? [] : [index]));
	if (options.scope === 'current') {
		const index = options.pageIndex ?? -1;
		return foreground.includes(index) ? [index] : [];
	}
	const from = Math.max(1, Math.floor(options.from ?? 1));
	const to = Math.min(foreground.length, Math.floor(options.to ?? foreground.length));
	return foreground.slice(from - 1, to).slice(0, MAX_PAGES);
}

/**
 * Pages as a PDF: vector paths, real (selectable) text and embedded pictures. Text uses the
 * 14 standard fonts (not embedded) unless the host passes `fonts`, which are subset and embedded.
 * Gradients are PDF shadings and shadows are offset copies; glow is a hard-edged halo; blur,
 * soft edges and reflections are not drawn. A page the vector path cannot represent (text no
 * supplied font covers, charts, patterns) becomes a picture page. The report says which and why
 * (`PdfPageReport.degraded` and `fallback`).
 */
export async function exportDocumentPdfDetailed(
	doc: Document,
	model: VisioDocument,
	options: PdfExportOptions = {},
): Promise<PdfExportResult> {
	const indexes = pdfPageIndexes(model, options);
	if (!indexes.length) throw new Error('The drawing has no foreground pages to export.');
	const fontSet = new PdfFontSet(options.fonts);
	const pages: PdfVectorPage[] = [];
	const reports: PdfPageReport[] = [];
	for (const index of indexes) {
		const source = model.pages[index]!;
		const exported = exportPageSvg(model, index);
		const parsed = new DOMParser().parseFromString(exported.svg, 'image/svg+xml').documentElement;
		const hrefs = collectSvgImages(parsed).slice(0, MAX_IMAGES_PER_PAGE);
		const images = new Map<string, PdfImageResource>();
		for (const href of hrefs) {
			const decoded = await decodeImage(doc, href);
			if (decoded) images.set(href, decoded);
		}
		const converted = convertSvgPage(parsed, exported.width, exported.height, images, fontSet);
		if (converted.page) {
			pages.push(converted.page);
			reports.push({
				index,
				name: source.name,
				mode: 'vector',
				fallback: [],
				degraded: converted.degraded,
			});
			continue;
		}
		const canvas = await rasterizePage(doc, model, index);
		const jpeg = await new Promise<Blob>((resolve, reject) =>
			canvas.toBlob(
				(blob) => (blob ? resolve(blob) : reject(new Error('picture'))),
				'image/jpeg',
				0.92,
			),
		);
		const page = new PdfVectorPage(exported.width * 72, exported.height * 72);
		page.save();
		page.concat([page.width, 0, 0, page.height, 0, 0]);
		page.drawImage({
			width: canvas.width,
			height: canvas.height,
			data: new Uint8Array(await jpeg.arrayBuffer()),
			filter: 'DCTDecode',
		});
		page.restore();
		pages.push(page);
		reports.push({
			index,
			name: source.name,
			mode: 'picture',
			fallback: converted.fallback,
			degraded: [],
		});
	}
	const bytes = buildVectorPdf(pages);
	const families = new Set((options.fonts ?? []).map((file) => file.family));
	const embedded = new Set(
		pages.flatMap((page) => [...page.embeddedFonts.values()].map((font) => font.family)),
	);
	return {
		blob: new Blob([bytes], { type: 'application/pdf' }),
		pages: reports,
		embeddedFonts: [...embedded],
		refusedFonts: [...families].filter(
			(family) => !fontSet.fonts.some((font) => font.family === family),
		),
	};
}

/** The PDF only; see `exportDocumentPdfDetailed` for the per-page report. */
export async function exportDocumentPdf(
	doc: Document,
	model: VisioDocument,
	options: PdfExportOptions = {},
): Promise<Blob> {
	return (await exportDocumentPdfDetailed(doc, model, options)).blob;
}
