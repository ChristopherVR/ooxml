import type { VisioDocument } from 'ooxml-core/visio';
import { PdfFontSet, type PdfFontFile } from 'ooxml-core/pdf/write';
import { buildVectorPdf, PdfVectorPage, type PdfImageResource } from 'ooxml-core/pdf/write';
import { exportPageSvg } from './export-svg';
import { collectSvgImages, convertSvgPage } from 'ooxml-core/pdf/write';
import { rasterizePage } from './viewer-export';
import { drawingFontFamily } from './font-stack';
import { decodePdfImage } from '../export/pdf-image';

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

const MAX_PAGES = 2000;
const MAX_IMAGES_PER_PAGE = 500;

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
			const decoded = await decodePdfImage(doc, href);
			if (decoded) images.set(href, decoded);
		}
		const converted = convertSvgPage(
			parsed,
			exported.width,
			exported.height,
			images,
			fontSet,
			drawingFontFamily,
		);
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
