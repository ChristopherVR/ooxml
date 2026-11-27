/**
 * Save a presentation as a vector PDF, the way every binding's PDF export calls it: each
 * selected slide is mounted by the binding (`mountSlide`, the same capture stage its picture
 * export uses), read into a display list and painted as paths, pictures and real text with the
 * deck's own embedded fonts subset into the file. Effects the PDF cannot express are pictures
 * of just those elements; a slide that cannot be read at all becomes one picture with its text
 * kept searchable. The result lists what was approximated, as codes a binding translates
 * (`pdf-export-diagnostics.ts`, `pdf-export-notice.ts`): the file is close to the slide, not
 * identical to it.
 */
import { collectLocalFontFiles } from '../../export/pdf-font-faces';
import { PdfFontSet, type PdfFontFile } from 'ooxml-core/pdf/write';
import { buildVectorPdf } from 'ooxml-core/pdf/write';
import type { SlideSizeEmu } from '../render/slide-size';
import type { CanvasSize } from '../types';
import { distinctPdfDiagnostics, type PdfExportDiagnostic } from './pdf-export-diagnostics';
import { collectSlide } from './pdf-vector-collect';
import type { DisplayItem } from './pdf-vector-display-list';
import { IDENTITY } from './pdf-vector-display-list';
import { paintSlide } from './pdf-vector-paint';
import { rasterizeIsolated } from './pdf-vector-raster';
import { externalFontCache } from './pdf-vector-raster-batch';
import type { PdfSlideSelection } from './pdf-vector-select';
import { selectPdfSlides, slidePdfPageSize } from './pdf-vector-select';
import type { RasterStrategyOptions } from './rasterize-element-strategy';

export interface VectorPdfExportOptions extends PdfSlideSelection {
	/** The document the stages are mounted in (fonts are read from its `@font-face` rules). */
	document: Document;
	/** The deck's slides (only `hidden` is read). */
	slides: readonly { hidden?: boolean }[];
	canvasSize: CanvasSize;
	/** `p:sldSz`, for an exact page size; the canvas size otherwise. */
	slideSize?: Pick<SlideSizeEmu, 'widthEmu' | 'heightEmu'>;
	/**
	 * Mount slide `index` laid out at the canvas size and resolve to its stage element once it
	 * has painted; `null` skips the slide. The stage stays mounted until the next call.
	 */
	mountSlide: (index: number) => Promise<HTMLElement | MountedPdfSlide | null | undefined>;
	/**
	 * Extra TrueType fonts to embed (subset). The deck's own embedded fonts are found in the
	 * page's `@font-face` rules; nothing is downloaded for the PDF.
	 */
	fonts?: readonly PdfFontFile[];
	/** Document title for the PDF's Info dictionary. */
	title?: string;
	/** Device pixels per CSS px for effects drawn as pictures. Default 2. */
	rasterScale?: number;
	/** Editor chrome the binding marks its own way (beyond `data-export-ignore`). */
	ignoreElement?: (element: Element) => boolean;
	/** The binding's html2canvas driver, the last resort for a picture capture. */
	html2canvasFallback?: RasterStrategyOptions['html2canvasFallback'];
	signal?: AbortSignal;
	onProgress?: (done: number, total: number) => void;
}

/** A mounted capture stage with the binding's html2canvas driver for it. */
export interface MountedPdfSlide {
	stage: HTMLElement;
	html2canvasFallback?: RasterStrategyOptions['html2canvasFallback'];
}

export interface VectorPdfSlideReport {
	/** 0-based slide index. */
	index: number;
	/** `vector` is paths and text (with any effects as pictures); `picture` is one image. */
	mode: 'vector' | 'picture';
	/** Elements of a `vector` slide drawn as pictures of just themselves. */
	pictures: number;
	diagnostics: PdfExportDiagnostic[];
}

export interface VectorPdfExportResult {
	blob: Blob;
	bytes: Uint8Array<ArrayBuffer>;
	slides: VectorPdfSlideReport[];
	/** Every distinct diagnostic of every slide, for one notice. */
	diagnostics: PdfExportDiagnostic[];
	/** Families embedded as subsets. */
	embeddedFonts: string[];
}

const abortError = () => new DOMException('Export cancelled', 'AbortError');

/** Export the selected slides as a vector PDF (see the module comment). */
export async function exportSlidesVectorPdf(
	options: VectorPdfExportOptions,
): Promise<VectorPdfExportResult> {
	const indexes = selectPdfSlides(options.slides, options);
	if (!indexes.length) throw new Error('There are no slides to export.');
	const { width: pageWidth, height: pageHeight } = slidePdfPageSize(
		options.canvasSize,
		options.slideSize,
	);
	const localFonts = await collectLocalFontFiles(options.document);
	const fonts = new PdfFontSet([...(options.fonts ?? []), ...localFonts]);
	const W = options.canvasSize.width,
		H = options.canvasSize.height;
	const externalFonts = externalFontCache(options.document);
	const pages = [];
	const reports: VectorPdfSlideReport[] = [];
	for (const [done, index] of indexes.entries()) {
		if (options.signal?.aborted) throw abortError();
		options.onProgress?.(done, indexes.length);
		const mounted = await options.mountSlide(index);
		if (!mounted) continue;
		const stage = 'stage' in mounted ? mounted.stage : mounted;
		const html2canvasFallback =
			('stage' in mounted ? mounted.html2canvasFallback : undefined) ?? options.html2canvasFallback;
		const collected = await collectSlide(stage, {
			canvasWidth: W,
			canvasHeight: H,
			fonts,
			externalFonts,
			...(options.rasterScale ? { rasterScale: options.rasterScale } : {}),
			...(html2canvasFallback ? { html2canvasFallback } : {}),
			...(options.ignoreElement ? { ignoreElement: options.ignoreElement } : {}),
		});
		let items: DisplayItem[] = collected.items;
		let mode: VectorPdfSlideReport['mode'] = 'vector';
		if (collected.picture) {
			mode = 'picture';
			const image = await rasterizeIsolated(
				stage,
				stage,
				'subtree',
				{ x: 0, y: 0, width: stage.offsetWidth || W, height: stage.offsetHeight || H },
				stage.offsetWidth || W,
				stage.offsetHeight || H,
				{
					scale: ((options.rasterScale ?? 2) * W) / (stage.offsetWidth || W),
					...(html2canvasFallback ? { html2canvasFallback } : {}),
					externalFonts,
				},
			);
			items = image
				? [
						{
							kind: 'image',
							matrix: IDENTITY,
							clips: [],
							alpha: 1,
							x: 0,
							y: 0,
							width: W,
							height: H,
							image,
						},
						...items,
					]
				: items;
			if (!image) collected.diagnostics.push({ code: 'slideNotCaptured' });
		}
		const painted = paintSlide(items, {
			canvasWidth: W,
			canvasHeight: H,
			pageWidth,
			pageHeight,
			fonts,
		});
		pages.push(painted.page);
		reports.push({
			index,
			mode,
			pictures: mode === 'vector' ? collected.pictures : 0,
			diagnostics: distinctPdfDiagnostics([...collected.diagnostics, ...painted.diagnostics]),
		});
	}
	if (!pages.length) throw new Error('There are no slides to export.');
	options.onProgress?.(indexes.length, indexes.length);
	const bytes = buildVectorPdf(pages, options.title ? { title: options.title } : {});
	const embedded = new Set(
		pages.flatMap((page) =>
			[...page.embeddedFonts.values()]
				.filter((font) => font !== fonts.textLayer)
				.map((font) => font.family),
		),
	);
	return {
		blob: new Blob([bytes], { type: 'application/pdf' }),
		bytes,
		slides: reports,
		diagnostics: distinctPdfDiagnostics(reports.flatMap((report) => report.diagnostics)),
		embeddedFonts: [...embedded],
	};
}
