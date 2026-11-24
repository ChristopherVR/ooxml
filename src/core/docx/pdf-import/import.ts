import {
	openPdf,
	type PdfLimits,
	type PdfPageContent,
	type PdfPageRule,
	type PdfTextSpan,
} from '../../pdf/index';
import { applyInternalLinks, applyOutline, resolveAnchors } from './anchors';
import { applyLinks, applyRules, pageLines, type Line } from './lines';
import { assemble } from './assemble';
import { attachFloats } from './floats';
import { extractFootnotes } from './footnotes';
import { applyLeaders, leaderTabs } from './leaders';
import { writePackage, type PictureBlock, type StoryContent } from './package';
import { buildParagraphs, rankHeadings, type ParagraphBlock, type TextArea } from './paragraphs';
import type { PdfImportReport, PdfOcrProvider } from './report';
import { rotatedText, type FloatingText } from './rotated';
import { extractStories } from './stories';
import { findGrids } from './table-grid';
import { ruledTables, type TableBlock } from './tables';
import { assignLists } from './lists';

export interface ImportPdfOptions {
	/** Pages to convert, 1-based and inclusive; all by default. */
	pages?: 'all' | { from: number; to: number };
	/** The password of an encrypted PDF (an empty user password opens without one). */
	password?: string;
	/** Detect running heads and feet and make them the header and footer (on by default). */
	headersFooters?: boolean;
	/** OCR for pages that are only a picture (none by default: such pages stay pictures). */
	ocr?: PdfOcrProvider;
	limits?: PdfLimits;
	signal?: AbortSignal;
}

/** The notice every converted document carries, as Word shows one. */
export const PDF_IMPORT_NOTICE =
	'This document was converted from a PDF. It may not look exactly like the original, especially if the PDF had many pictures or tables. Saving writes a Word document; the PDF itself is unchanged.';

const median = (values: number[], fallback: number) => {
	if (!values.length) return fallback;
	const sorted = [...values].sort((a, b) => a - b);
	return sorted[Math.floor(sorted.length / 2)]!;
};

function weightedMode<T>(entries: [T, number][], fallback: T): T {
	const weights = new Map<T, number>();
	for (const [value, weight] of entries) weights.set(value, (weights.get(value) ?? 0) + weight);
	let best = fallback;
	let max = 0;
	for (const [value, weight] of weights) if (weight > max) [best, max] = [value, weight];
	return best;
}

function ocrSpans(
	words: { text: string; x: number; y: number; width: number; height: number }[],
): PdfTextSpan[] {
	return words
		.filter((word) => word.text.trim() && word.width > 0 && word.height > 0)
		.map((word) => ({
			text: word.text,
			x: word.x,
			y: word.y + word.height * 0.8,
			width: word.width,
			fontSize: word.height * 0.9,
			font: {
				name: 'Calibri',
				family: 'Calibri',
				bold: false,
				italic: false,
				serif: false,
				monospace: false,
			},
			color: '#000000',
			angle: 0,
			rise: 0,
		}));
}

interface PageWork {
	index: number;
	width: number;
	height: number;
	lines: Line[];
	pictures: (PictureBlock & { x: number })[];
	tables: TableBlock[];
	notes: string[];
	rules: PdfPageRule[];
	rotated: FloatingText[];
	notesBottom?: number;
	scanned?: true;
}

async function readPage(
	content: PdfPageContent,
	ocr: PdfOcrProvider | undefined,
): Promise<PageWork> {
	const notes = [...content.notes];
	let spans = content.spans;
	const visible = spans.filter((span) => span.text.trim() && !span.invisible);
	const pictures = content.images.filter(
		(image): image is typeof image & { data: Uint8Array; format: 'png' | 'jpeg' } =>
			image.format !== 'unsupported' && !!image.data && image.width > 6 && image.height > 6,
	);
	for (const image of content.images)
		if (image.format === 'unsupported' && image.width > 6 && image.height > 6)
			notes.push(`A picture was left out: ${image.reason ?? 'it could not be read'}.`);
	const fullPage = pictures.find(
		(image) => image.width * image.height > content.width * content.height * 0.5,
	);
	let scanned: true | undefined;
	if (!visible.length && fullPage) {
		scanned = true;
		if (spans.length)
			notes.push(
				'The page is a scan with a text layer: its text was converted and the scan left out.',
			);
		else if (ocr) {
			spans = ocrSpans(
				await ocr({
					index: content.index,
					width: content.width,
					height: content.height,
					image: {
						data: fullPage.data,
						format: fullPage.format,
						x: fullPage.x,
						y: fullPage.y,
						width: fullPage.width,
						height: fullPage.height,
					},
				}),
			);
			notes.push('The page is a scan: its text comes from OCR and the scan was left out.');
		} else notes.push('The page is a scan without text: it was kept as a picture.');
	}
	const { lines: all } = pageLines(content.index, spans);
	const rotated = rotatedText(content.index, spans, 11);
	if (rotated.lines)
		notes.push(
			`Rotated text (${rotated.lines} line${rotated.lines === 1 ? '' : 's'}) was put in text boxes turned like the original.`,
		);
	// Table rules are not underlines: only the rules no table uses mark text.
	const { grids, rest } = findGrids(content.rules);
	applyRules(all, rest);
	applyLinks(all, content.links);
	applyInternalLinks(all, content.links);
	applyLeaders(all);
	const { tables, lines } = ruledTables(content.index, all, grids, 11);
	const keep = scanned && spans.length ? pictures.filter((image) => image !== fullPage) : pictures;
	return {
		index: content.index,
		width: content.width,
		height: content.height,
		lines,
		pictures: keep.map((image) => ({
			kind: 'picture',
			page: content.index,
			top: image.y,
			data: image.data,
			format: image.format,
			width: image.width,
			height: image.height,
			align: 'left',
			x: image.x,
			...(image.alt ? { alt: image.alt } : {}),
		})),
		tables,
		notes,
		rules: content.rules,
		rotated: rotated.floats,
		...(scanned ? { scanned } : {}),
	};
}

export interface ConvertPdfResult {
	/** The converted document as .docx bytes. */
	docx: Uint8Array;
	report: PdfImportReport;
	/** The notices a document opened from it shows first (`PDF_IMPORT_NOTICE` and the report's notes). */
	warnings: string[];
}

/**
 * Converts `bytes` (a PDF) to a Word document's package. A host that already loads .docx files
 * opens `docx` the usual way; `importPdf` does both.
 */
export async function convertPdf(
	bytes: Uint8Array,
	options: ImportPdfOptions = {},
): Promise<ConvertPdfResult> {
	const reader = await openPdf(bytes, {
		...(options.password !== undefined ? { password: options.password } : {}),
		...(options.limits ? { limits: options.limits } : {}),
	});
	const range =
		options.pages && options.pages !== 'all' ? options.pages : { from: 1, to: reader.pageCount };
	const pages: PageWork[] = [];
	for (
		let index = Math.max(0, range.from - 1);
		index < Math.min(reader.pageCount, range.to);
		index++
	) {
		options.signal?.throwIfAborted();
		pages.push(await readPage(await reader.page(index), options.ocr));
	}
	if (!pages.length) throw new RangeError('The page range selects no page of the PDF.');
	const repeated = options.headersFooters === false ? [] : extractStories(pages);
	// Table text counts for the body size, the fonts and the text area like any other text.
	const tableLines = (page: PageWork) =>
		page.tables.flatMap((table) =>
			table.cells.flatMap((cell) => cell.blocks.flatMap((block) => block.lines)),
		);
	const allLines = pages.flatMap((page) => [...page.lines, ...tableLines(page)]);
	const extents = pages
		.map((page) => [
			...page.lines,
			...page.tables.map((table) => ({
				left: table.left,
				right: table.left + table.columns.reduce((sum, width) => sum + width, 0),
			})),
		])
		.filter((boxes) => boxes.length);
	const bodySize = weightedMode(
		allLines.map((line) => [Math.round(line.size * 2) / 2, line.text.length]),
		11,
	);
	const family = weightedMode(
		allLines.flatMap((line) =>
			line.fragments.map(
				(fragment) => [fragment.font.family, fragment.text.length] as [string, number],
			),
		),
		'Calibri',
	);
	const footnotes = extractFootnotes(pages, bodySize);
	const area: TextArea = {
		left: median(
			extents.map((boxes) => Math.min(...boxes.map((box) => box.left))),
			72,
		),
		right: median(
			extents.map((boxes) => Math.max(...boxes.map((box) => box.right))),
			540,
		),
	};
	const { blocks, sections } = assemble(pages, area, bodySize, repeated);
	// Paragraphs in table cells count too, for links, bookmarks, headings and leaders.
	const paragraphBlocks = blocks.flatMap((block): ParagraphBlock[] =>
		block.kind === 'paragraph'
			? [block]
			: block.kind === 'table'
				? block.cells.flatMap((cell) => cell.blocks)
				: [],
	);
	rankHeadings(blocks.filter((block): block is ParagraphBlock => block.kind === 'paragraph'));
	const numbering = assignLists(blocks);
	applyOutline(paragraphBlocks, reader.outline);
	leaderTabs(paragraphBlocks, area);
	const unresolved = resolveAnchors(paragraphBlocks);
	attachFloats(
		blocks,
		sections,
		pages.flatMap((page) => page.rotated),
		bodySize,
	);
	const noteBlocks = footnotes.map((note) => ({
		id: note.id,
		...(note.customMark !== undefined ? { customMark: note.customMark } : {}),
		blocks: buildParagraphs(note.lines, area, bodySize).map(({ heading: _heading, ...block }) => ({
			...block,
			indentLeft: 0,
			firstLine: 0,
			spaceBefore: 0,
		})),
	}));
	const stories: StoryContent[] = repeated.map((story) => ({
		kind: story.kind,
		blocks: buildParagraphs(story.lines, area, bodySize).map(({ heading: _heading, ...block }) => ({
			...block,
			spaceBefore: 0,
		})),
		pageNumber: (story.lines[0]?.page ?? 0) + 1,
		pageCount: reader.pageCount,
	}));
	const info = reader.info;
	const packageBytes = await writePackage({
		blocks,
		numbering,
		sections,
		stories,
		defaults: { family, size: bodySize },
		properties: {
			...(info.title ? { title: info.title } : {}),
			...(info.author ? { author: info.author } : {}),
			...(info.subject ? { subject: info.subject } : {}),
			...(info.keywords ? { keywords: info.keywords } : {}),
		},
		...(info.language ? { language: info.language } : {}),
		...(noteBlocks.length ? { footnotes: noteBlocks } : {}),
	});
	const notes: string[] = [];
	if (unresolved)
		notes.push(
			unresolved === 1
				? 'An internal link points to text that was not converted; it was left as plain text.'
				: `${unresolved} internal links point to text that was not converted; they were left as plain text.`,
		);
	if (reader.signed)
		notes.push('The PDF was digitally signed; the converted document has no signature.');
	if (pages.some((page) => page.scanned))
		notes.push('Some pages are scans; see the per-page notes.');
	const report: PdfImportReport = {
		pages: pages.map((page) => ({
			index: page.index,
			notes: page.notes,
			...(page.scanned ? { scanned: true as const } : {}),
		})),
		fonts: [
			...new Set(
				allLines.flatMap((line) => line.fragments.map((fragment) => fragment.font.family)),
			),
		].sort(),
		encrypted: reader.encrypted,
		...(reader.permissions ? { permissions: reader.permissions } : {}),
		signed: reader.signed,
		notes,
	};
	return { docx: packageBytes, report, warnings: [PDF_IMPORT_NOTICE, ...notes] };
}
