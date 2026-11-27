/**
 * Saves a Word document as a real PDF, the way Word's Save As PDF does: selectable text in the
 * document's fonts (embedded as subsets when the host supplies them), vector borders, shading and
 * rules, charts, shapes and text boxes as vector drawings with real text, pictures at their own
 * resolution, watermarks, page borders and line numbers, on the pages Print Layout shows. It paints the
 * layout engine's result, so line and page breaks match the screen. No page is rasterised and no
 * print dialog opens. DOM-free: give a measurer to run it without a canvas.
 */
import {
	refreshComputedFields,
	reviewDocumentFormatting,
	reviewFormattingWarnings,
	type DocumentModel,
	type ReviewDisplayMode,
} from '../index';
import {
	createFakeMeasurer,
	createPageStories,
	type PlacedStory,
	layoutDocumentModel,
	type LayoutPageBox,
	type OversizedCantSplitRow,
	type TextMeasurer,
} from '../layout/index';
import { PdfFontSet, type PdfFontFile } from '../../pdf/write/pdf-embedded-font';
import { buildVectorPdf, PdfVectorPage, type PdfVectorOptions } from '../../pdf/write/pdf-vector';
import { DocxPdfFontFiles } from './font-files';
import { createFontFileMeasurer, type FontFileMeasurer } from './font-measurer';
import { DocxPdfBalloons, type DocxPdfMarkupBalloons } from './pdf-balloon-items';
import { paintBlock, type DocxPdfPainter } from './pdf-blocks';
import { DocxPdfDrawings } from './pdf-drawings';
import { paintFootnotes } from './pdf-footnotes';
import { DocxPdfHooks } from './pdf-hooks';
import type { DocxPdfBookmarks } from './pdf-navigation';
import { blockPictures, pagePictures } from './pdf-objects';
import { pageMarks, paintPageMarks, type DocxPdfPageMarks } from './pdf-page-marks';
import { DocxPdfPictures, type DocxPdfMedia, type DocxPdfPictureDecoder } from './pdf-pictures';
import {
	DocxPdfReportBuilder,
	pdfAConformance,
	pdfUaConformance,
	type DocxPdfReport,
} from './pdf-report';
import {
	paintStoryFloats,
	storyGraphicPictures,
	storyGraphics,
	type DocxPdfStoryGraphics,
} from './pdf-story-graphics';
import { DocxPdfText } from './pdf-text';
import { paintWatermark } from './pdf-watermark';

export type {
	DocxPdfFontMeasurement,
	DocxPdfReport,
	DocxPdfOmittedKind,
	DocxPdfConformance,
	DocxPdfUaConformance,
} from './pdf-report';
export type { DocxPdfBookmarks } from './pdf-navigation';
export type { DocxPdfMarkupBalloons } from './pdf-balloon-items';
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
	 * Word's "Balloons" setting, with `markup: 'all'`. Pages with balloons are shrunk to leave a
	 * markup pane on the right, as Word prints them; each balloon is joined to its text by a
	 * dotted line, and comments stay PDF notes too. `commentsAndFormatting` (Word's default) shows
	 * comments and formatting changes ("Formatted: Font: Bold") in balloons and insertions and
	 * deletions in the text; `all` moves deletions to balloons ("Deleted: ..."); `comments` shows
	 * only comments in balloons (Word's "Show all revisions inline"); `none` draws no pane.
	 */
	markupBalloons?: DocxPdfMarkupBalloons;
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
	 * paragraphs, lists, tables, figures (with their alternative text) and links read in order, with
	 * headers and footers as artifacts. On by default.
	 */
	taggedPdf?: boolean;
	/**
	 * Word's "PDF/A compliant": writes PDF/A-2 metadata (XMP), an sRGB output intent and the file
	 * identifier, and writes text in families no supplied font is named for with a supplied font
	 * that covers it (as a substitution) rather than an unembedded standard font. The file only
	 * conforms when every font is embedded, and then declares PDF/A-2u; `report.pdfA` says whether
	 * it does and at which level (see `docs/docx-pdf-export.md`, validated with veraPDF).
	 */
	pdfA?: boolean;
	/**
	 * Requested PDF/A-2 conformance level when `pdfA` is true: `'U'` (default) or `'A'` (requires
	 * structure tags, document language and alternative text on every figure). Only declared when
	 * every requirement is met.
	 */
	pdfALevel?: 'A' | 'U';
	/**
	 * Requests PDF/UA-1 (ISO 14289-1) accessibility metadata (`pdfuaid:part` = 1) and checks
	 * structure tags, embedded fonts, document title, language and figure alternative text.
	 * Declared only when every requirement is met; `report.pdfUa` reports conformance and issues.
	 */
	pdfUa?: boolean;
	/** Natural language of the document, e.g. `en-US`. */
	language?: string;
	/**
	 * Font files to embed (TrueType). Text in a family with no file is written with the supplied
	 * file `fontSubstitutions` or the substitution table names for it (Carlito for Calibri,
	 * Liberation Serif for Times New Roman...), else with a standard PDF font. Without a
	 * `measurer`, the layout measures with these files (see `measurer`).
	 */
	fonts?: readonly PdfFontFile[];
	/**
	 * The host's font mapping: a document family to the supplied family that stands in for it
	 * (`{ 'Segoe UI': 'Noto Sans' }`), tried before the built-in substitution table.
	 */
	fontSubstitutions?: Readonly<Record<string, string>>;
	/** The document's picture parts (as `loadDocx` returns them, or a lookup). */
	media?: DocxPdfMedia;
	/** Decodes picture formats other than JPEG and PNG (GIF, BMP, TIFF...) to RGBA. */
	decodePicture?: DocxPdfPictureDecoder;
	/**
	 * Measures text for the layout. By default the supplied `fonts` measure it
	 * (`createFontFileMeasurer`: their advance widths, Word's kerning and line heights; families with
	 * no file are estimated and listed in `report.fontMeasurement`), and without fonts a fixed
	 * formula does. `ooxml-ui/docx`'s `exportDocxPdf` passes the canvas measurer Print Layout uses.
	 */
	measurer?: TextMeasurer;
	/** How to paginate a `w:cantSplit` row taller than an empty page/column (`'split'` by default). */
	oversizedCantSplitRow?: OversizedCantSplitRow;
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

function selectedPages(count: number, pages: DocxPdfOptions['pages']): number[] {
	const all = Array.from({ length: count }, (_, index) => index);
	if (!pages || pages === 'all') return all;
	const from = Math.max(1, Math.floor(pages.from));
	const to = Math.min(count, Math.floor(pages.to));
	return all.filter((index) => index + 1 >= from && index + 1 <= to);
}

/** What a page draws besides its text flow (`pdf-page-marks.ts`, `pdf-story-graphics.ts`). */
interface PageExtras {
	marks: DocxPdfPageMarks | undefined;
	graphics: DocxPdfStoryGraphics | undefined;
	measurer: TextMeasurer;
}

/** Paints one laid-out page. */
function paintPage(
	painter: DocxPdfPainter,
	page: LayoutPageBox,
	stories: readonly PlacedStory[],
	model: DocumentModel,
	report: DocxPdfReportBuilder,
	{ marks, graphics, measurer }: PageExtras,
): void {
	const pdf = painter.page;
	pdf.concat([PT_PER_PX, 0, 0, -PT_PER_PX, 0, page.heightPx * PT_PER_PX]);
	const hooks = painter.hooks;
	hooks?.startPage(pdf, page);
	// Page borders, line numbers, watermarks and floats behind the header and footer are
	// pagination artifacts in a tagged PDF (PDF/UA: all content is tagged or an artifact).
	const decoration = (paint: () => void) => (hooks ? hooks.tagger.asArtifact(paint) : paint());
	const marked = !!marks?.borders || !!marks?.lineNumbers.length;
	if (
		marked ||
		graphics?.watermark ||
		graphics?.pictureWatermark ||
		graphics?.floats.some((float) => float.behindText)
	)
		decoration(() => {
			paintPageMarks(painter, page, marks, 'back', model, measurer, report);
			paintWatermark(painter, graphics, page, measurer);
			paintStoryFloats(painter, graphics, 'header', true);
			paintStoryFloats(painter, graphics, 'footer', true);
		});
	const floats = page.floats ?? [];
	const paintFloat = (float: (typeof floats)[number]) =>
		painter.pictures.draw(pdf, float, float.xPx, float.yPx);
	const placeFloat = (float: (typeof floats)[number]) =>
		hooks ? hooks.float(float, () => paintFloat(float)) : paintFloat(float);
	for (const float of floats) if (float.behindText) placeFloat(float);
	// Header, body, notes, footer: the order text is read and extracted in. Headers and footers
	// are pagination artifacts in a tagged PDF.
	const paintStories = (kind: 'header' | 'footer') => {
		for (const story of stories)
			if (story.kind === kind) {
				hooks?.story(story.source);
				for (const block of story.blocks) paintBlock(painter, block, story.x, story.y);
				hooks?.story(undefined);
			}
		paintStoryFloats(painter, graphics, kind, false);
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
	if (marked)
		decoration(() => paintPageMarks(painter, page, marks, 'front', model, measurer, report));
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
	const files = new DocxPdfFontFiles(options.fonts ?? [], options.fontSubstitutions);
	const fileMeasurer: FontFileMeasurer | undefined =
		options.measurer || !files.faces.length ? undefined : createFontFileMeasurer(files);
	const measurer = options.measurer ?? fileMeasurer ?? createFakeMeasurer();
	const report = new DocxPdfReportBuilder();
	const projected = reviewDocumentFormatting(refreshComputedFields(model), mode);
	const balloonMode = options.markupBalloons ?? 'commentsAndFormatting';
	// Deletions shown in balloons leave the text, as Word's "Show revisions in balloons" does.
	const layout = layoutDocumentModel(projected.model, measurer, {
		reviewDisplayMode: mode === 'all' && balloonMode === 'all' ? 'simple' : mode,
		...(options.oversizedCantSplitRow
			? { oversizedCantSplitRow: options.oversizedCantSplitRow }
			: {}),
	});
	for (const message of [
		...layout.approximations,
		...reviewFormattingWarnings(projected.diagnostics),
	])
		report.note(message);
	const indices = selectedPages(layout.pages.length, options.pages);
	if (!indices.length) throw new RangeError('The page range selects no page of the document.');
	const placeStories = createPageStories(
		projected.model,
		layout.pages,
		measurer,
		{ reviewDisplayMode: mode },
		(message) => report.note(message),
		options.now,
	);
	const stories = layout.pages.map((_, index) => placeStories(index));
	const marks = pageMarks(projected.model, layout.pages);
	const graphics = storyGraphics(stories, mode);
	const text = new DocxPdfText(
		new PdfFontSet(options.fonts ?? [], {
			embedAll: !!(options.pdfA || options.pdfUa),
			substitute: (family) => files.match(family)?.family,
		}),
		measurer,
		report,
	);
	const pictures = new DocxPdfPictures(options.media, options.decodePicture, report);
	pictures.drawings = new DocxPdfDrawings(pictures, text, report);
	for (const index of indices) {
		await pictures.prepare(pagePictures(layout.pages[index]!));
		for (const story of stories[index] ?? []) await pictures.prepare(blockPictures(story.blocks));
		if (graphics[index]) await pictures.prepare(storyGraphicPictures(graphics[index]));
	}
	const tagged = options.taggedPdf !== false;
	const hooks = new DocxPdfHooks(projected.model, {
		tagged,
		bookmarks: options.bookmarks ?? 'headings',
		mode,
		balloons: new DocxPdfBalloons(projected.model, balloonMode, mode === 'all', text, report),
	});
	pictures.drawings.hooks = hooks;
	const pages: PdfVectorPage[] = [];
	for (const index of indices) {
		signal?.throwIfAborted();
		const page = layout.pages[index]!;
		report.page = index + 1;
		const pdf = new PdfVectorPage(page.widthPx * PT_PER_PX, page.heightPx * PT_PER_PX);
		const shrink = hooks.options.balloons?.frame(page);
		if (shrink) pdf.concat(shrink.matrix);
		paintPage(
			{ page: pdf, text, pictures, hooks },
			page,
			stories[index] ?? [],
			projected.model,
			report,
			{
				marks: marks[index],
				graphics: graphics[index],
				measurer,
			},
		);
		pages.push(pdf);
	}
	const { outline, unresolved } = hooks.navigation.finish();
	hooks.tagger.finish();
	if (unresolved.length)
		report.note(
			`Links to bookmarks on no page written stay plain text: ${unresolved.slice(0, 5).join(', ')}${unresolved.length > 5 ? '...' : ''}.`,
		);
	if (fileMeasurer)
		report.measured(
			[...fileMeasurer.families].map(([family, match]) =>
				match
					? { family, font: match.family, source: match.source }
					: { family, font: null, source: 'estimated' as const },
			),
		);
	const properties = options.documentProperties === false ? undefined : model.properties;
	const built = report.build(pages.length);
	const check = {
		...(options.pdfALevel ? { requestedLevel: options.pdfALevel } : {}),
		tagged,
		...(tagged ? { structure: hooks.tagger.root } : {}),
		...(options.language ? { language: options.language } : {}),
		...(properties?.title ? { title: properties.title } : {}),
	};
	if (options.pdfA) built.pdfA = pdfAConformance(built, pages, check);
	if (options.pdfUa) built.pdfUa = pdfUaConformance(built, pages, check);
	const info: PdfVectorOptions = {
		...(properties?.title ? { title: properties.title } : {}),
		...(properties?.creator ? { author: properties.creator } : {}),
		...(properties?.subject ? { subject: properties.subject } : {}),
		...(properties?.keywords ? { keywords: properties.keywords } : {}),
		...(options.language ? { language: options.language } : {}),
		...(tagged ? { structure: hooks.tagger.root } : {}),
		...(outline.length ? { outline } : {}),
		...(options.pdfA
			? {
					pdfA: true,
					...(built.pdfA ? { pdfALevel: built.pdfA.level } : {}),
					...(options.now ? { date: options.now } : {}),
				}
			: {}),
		...(options.pdfUa && built.pdfUa?.conformant
			? { pdfUa: true, ...(options.now ? { date: options.now } : {}) }
			: {}),
	};
	const bytes = buildVectorPdf(pages, info);
	return { bytes, blob: new Blob([bytes], { type: 'application/pdf' }), report: built };
}
