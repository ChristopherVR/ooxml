/**
 * Saves a Word document as a real PDF, the way Word's Save As PDF does: selectable text in the
 * document's fonts (embedded as subsets when the host supplies them), vector borders, shading and
 * rules, and pictures at their own resolution, on the pages Print Layout shows. It paints the
 * layout engine's result, so line and page breaks match the screen. No page is rasterised and no
 * print dialog opens. DOM-free: give a measurer to run it without a canvas.
 */
import {
	refreshComputedFields,
	reviewDocumentFormatting,
	reviewFormattingWarnings,
	type DocumentModel,
	type ReviewDisplayMode,
} from 'ooxml-core/docx';
import {
	layoutDocumentModel,
	type LayoutBlockBox,
	type LayoutPageBox,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import { PdfFontSet, type PdfFontFile } from '../../export/pdf-embedded-font';
import { buildVectorPdf, PdfVectorPage, type PdfVectorOptions } from '../../export/pdf-vector';
import { createCanvasMeasurer } from '../canvas-measurer';
import { paintBlock, type DocxPdfPainter } from './pdf-blocks';
import { paintFootnotes } from './pdf-footnotes';
import { DocxPdfHooks } from './pdf-hooks';
import type { DocxPdfBookmarks } from './pdf-navigation';
import {
	DocxPdfPictures,
	type DocxPdfMedia,
	type DocxPdfPicture,
	type DocxPdfPictureDecoder,
} from './pdf-pictures';
import { DocxPdfReportBuilder, pdfAConformance, type DocxPdfReport } from './pdf-report';
import { placeStories } from './pdf-stories';
import { DocxPdfText } from './pdf-text';

export type { DocxPdfReport, DocxPdfOmittedKind, DocxPdfConformance } from './pdf-report';
export type { DocxPdfBookmarks } from './pdf-navigation';
export type { DecodedPicture, DocxPdfMedia, DocxPdfPictureDecoder } from './pdf-pictures';

export interface DocxPdfOptions {
	/** Pages to write, 1-based and inclusive; all by default. */
	pages?: 'all' | { from: number; to: number };
	/**
	 * `none` writes the document as it reads with changes accepted (Word's "Document"); `all`
	 * shows tracked changes as Print Layout's All Markup does (insertions underlined, deletions
	 * struck through in each author's colour, change bars) and comments as PDF notes.
	 */
	markup?: 'none' | 'all';
	/**
	 * Word's "Create bookmarks using": `headings` (the default) turns headings into the PDF's
	 * outline, nested by level; `wordBookmarks` lists the document's own bookmarks (not hidden
	 * ones such as `_Toc`); `none` writes no outline.
	 */
	bookmarks?: DocxPdfBookmarks;
	/** Writes the title, author, subject and keywords into the PDF. On by default. */
	documentProperties?: boolean;
	/**
	 * Word's "Document structure tags for accessibility": a tagged PDF whose headings,
	 * paragraphs, tables, figures (with their alternative text) and links read in order, with
	 * headers and footers as artifacts. On by default.
	 */
	taggedPdf?: boolean;
	/**
	 * Word's "PDF/A compliant": writes PDF/A-2b metadata (XMP), an sRGB output intent and the file
	 * identifier. The file only conforms when every font is embedded; `report.pdfA` says whether
	 * it does.
	 */
	pdfA?: boolean;
	/** Natural language of the document, e.g. `en-US`. */
	language?: string;
	/** Font files to embed (TrueType); text in other fonts uses the standard PDF fonts. */
	fonts?: readonly PdfFontFile[];
	/** The document's picture parts (as `loadDocx` returns them, or a lookup). */
	media?: DocxPdfMedia;
	/** Decodes picture formats other than JPEG and PNG (GIF, BMP, TIFF...) to RGBA. */
	decodePicture?: DocxPdfPictureDecoder;
	/** Measures text for the layout; a canvas measurer by default (the one Print Layout uses). */
	measurer?: TextMeasurer;
	/** The time DATE and TIME fields show; now by default. */
	now?: Date;
	signal?: AbortSignal;
}

export interface DocxPdfResult {
	bytes: Uint8Array<ArrayBuffer>;
	blob: Blob;
	report: DocxPdfReport;
}

const PT_PER_PX = 0.75;

function* blockPictures(blocks: readonly LayoutBlockBox[]): Generator<DocxPdfPicture> {
	for (const block of blocks) {
		if (block.kind === 'paragraph') {
			for (const line of block.lines)
				for (const fragment of line.fragments) if (fragment.object) yield fragment.object;
		} else
			for (const row of block.rows) {
				for (const cell of row.cells) yield* blockPictures(cell);
				for (const nested of row.cellTables ?? []) yield* blockPictures(nested);
			}
	}
}

function* pagePictures(page: LayoutPageBox): Generator<DocxPdfPicture> {
	for (const column of page.columns) yield* blockPictures(column.blocks);
	for (const note of page.footnotes ?? []) yield* blockPictures(note.paragraphs);
	yield* page.floats ?? [];
}

function selectedPages(count: number, pages: DocxPdfOptions['pages']): number[] {
	const all = Array.from({ length: count }, (_, index) => index);
	if (!pages || pages === 'all') return all;
	const from = Math.max(1, Math.floor(pages.from));
	const to = Math.min(count, Math.floor(pages.to));
	return all.filter((index) => index + 1 >= from && index + 1 <= to);
}

/** Paints one laid-out page. */
function paintPage(
	painter: DocxPdfPainter,
	page: LayoutPageBox,
	stories: ReturnType<typeof placeStories>[number],
	model: DocumentModel,
	report: DocxPdfReportBuilder,
): void {
	const pdf = painter.page;
	pdf.concat([PT_PER_PX, 0, 0, -PT_PER_PX, 0, page.heightPx * PT_PER_PX]);
	const section = model.sections?.[page.sectionIndex];
	if (section?.pageBorders)
		report.omitted('pageBorder', 'Page borders are not exported to PDF yet.');
	if (section?.lineNumbering) report.note('Line numbers are not exported to PDF yet.');
	const floats = page.floats ?? [];
	const paintFloat = (float: (typeof floats)[number]) => {
		if (float.textBox || float.textBoxLayout)
			report.omitted('textBox', 'Text boxes are not exported to PDF yet.');
		painter.pictures.draw(pdf, float, float.xPx, float.yPx);
	};
	const hooks = painter.hooks;
	hooks?.startPage(pdf, page);
	const placeFloat = (float: (typeof floats)[number]) =>
		hooks ? hooks.float(float, () => paintFloat(float)) : paintFloat(float);
	for (const float of floats) if (float.behindText) placeFloat(float);
	// Header, body, notes, footer: the order text is read and extracted in. Headers and footers
	// are pagination artifacts in a tagged PDF.
	const paintStories = (kind: 'header' | 'footer') => {
		for (const story of stories)
			if (story.kind === kind)
				for (const block of story.blocks) paintBlock(painter, block, story.x, story.y);
	};
	const artifact = (paint: () => void, subtype: 'Header' | 'Footer') =>
		hooks && stories.some((story) => story.kind === subtype.toLowerCase())
			? hooks.tagger.asArtifact(paint, subtype)
			: paint();
	artifact(() => paintStories('header'), 'Header');
	for (const column of page.columns) {
		for (const block of column.blocks)
			paintBlock(painter, block, page.marginLeftPx + column.xPx, page.marginTopPx);
	}
	paintFootnotes(painter, page);
	artifact(() => paintStories('footer'), 'Footer');
	for (const float of floats) if (!float.behindText) placeFloat(float);
	hooks?.endPage();
}

/** Writes `model` as a PDF (see the module comment). */
export async function exportDocxPdf(
	model: DocumentModel,
	options: DocxPdfOptions = {},
): Promise<DocxPdfResult> {
	const signal = options.signal;
	signal?.throwIfAborted();
	const mode: ReviewDisplayMode = options.markup === 'all' ? 'all' : 'final';
	const measurer = options.measurer ?? createCanvasMeasurer();
	const report = new DocxPdfReportBuilder();
	const projected = reviewDocumentFormatting(refreshComputedFields(model), mode);
	const layout = layoutDocumentModel(projected.model, measurer, { reviewDisplayMode: mode });
	for (const message of [
		...layout.approximations,
		...reviewFormattingWarnings(projected.diagnostics),
	])
		report.note(message);
	const indices = selectedPages(layout.pages.length, options.pages);
	if (!indices.length) throw new RangeError('The page range selects no page of the document.');
	const stories = placeStories(
		projected.model,
		layout.pages,
		measurer,
		mode,
		report,
		options.now ?? new Date(),
	);
	const pictures = new DocxPdfPictures(options.media, options.decodePicture, report);
	for (const index of indices) {
		await pictures.prepare(pagePictures(layout.pages[index]!));
		for (const story of stories[index] ?? []) await pictures.prepare(blockPictures(story.blocks));
	}
	const text = new DocxPdfText(new PdfFontSet(options.fonts ?? []), measurer, report);
	const tagged = options.taggedPdf !== false;
	const hooks = new DocxPdfHooks(projected.model, {
		tagged,
		bookmarks: options.bookmarks ?? 'headings',
		mode,
	});
	const pages: PdfVectorPage[] = [];
	for (const index of indices) {
		signal?.throwIfAborted();
		const page = layout.pages[index]!;
		report.page = index + 1;
		const pdf = new PdfVectorPage(page.widthPx * PT_PER_PX, page.heightPx * PT_PER_PX);
		paintPage(
			{ page: pdf, text, pictures, hooks },
			page,
			stories[index] ?? [],
			projected.model,
			report,
		);
		pages.push(pdf);
	}
	const { outline, unresolved } = hooks.navigation.finish();
	if (unresolved.length)
		report.note(
			`Links to bookmarks on no page written stay plain text: ${unresolved.slice(0, 5).join(', ')}${unresolved.length > 5 ? '...' : ''}.`,
		);
	const properties = options.documentProperties === false ? undefined : model.properties;
	const info: PdfVectorOptions = {
		...(properties?.title ? { title: properties.title } : {}),
		...(properties?.creator ? { author: properties.creator } : {}),
		...(properties?.subject ? { subject: properties.subject } : {}),
		...(properties?.keywords ? { keywords: properties.keywords } : {}),
		...(options.language ? { language: options.language } : {}),
		...(tagged ? { structure: hooks.tagger.root } : {}),
		...(outline.length ? { outline } : {}),
		...(options.pdfA ? { pdfA: true, ...(options.now ? { date: options.now } : {}) } : {}),
	};
	const bytes = buildVectorPdf(pages, info);
	const built = report.build(pages.length);
	if (options.pdfA) built.pdfA = pdfAConformance(built, pages);
	return { bytes, blob: new Blob([bytes], { type: 'application/pdf' }), report: built };
}
