/**
 * Headers and footers for the Word PDF export. Print Layout draws them as HTML; the export lays
 * each page's header and footer out with the layout engine instead, at the width of the page's
 * text column, with that page's field results (PAGE, NUMPAGES, STYLEREF...) filled in, and
 * places them at the section's header and footer distances. DOM-free.
 */
import {
	isRunHiddenForReview,
	signedTwips,
	twips,
	type Block,
	type DocumentModel,
	type HeaderFooterContent,
	type Paragraph,
	type ReviewDisplayMode,
	type TextRun,
} from 'ooxml-core/docx';
import {
	fieldDisplayText,
	headerFooterForPage,
	layoutDocumentModel,
	pageNumberValues,
	pageNumbers,
	pageStyleRefs,
	sectionPageCounts,
	type LayoutBlockBox,
	type LayoutPageBox,
	type PageFieldContext,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import type { DocxPdfReportBuilder } from './pdf-report';

/** A header or footer laid out for one page, in page coordinates. */
export interface PlacedStory {
	kind: 'header' | 'footer';
	x: number;
	y: number;
	blocks: LayoutBlockBox[];
}

const TALL_PAGE_TWIPS = 31680 * 4;
const PX_TWIPS = 15;

function fillParagraph(
	paragraph: Paragraph,
	values: PageFieldContext,
	mode: ReviewDisplayMode,
	report: DocxPdfReportBuilder,
): Paragraph {
	const runs: TextRun[] = [];
	for (const run of paragraph.runs) {
		if (isRunHiddenForReview(run, mode)) continue;
		if (run.image?.watermark || run.image?.pictureWatermark) {
			report.omitted('watermark', 'Watermarks are not exported to PDF yet.');
			continue;
		}
		if (run.image?.anchored) {
			report.omitted(
				'headerFooterFloat',
				'Floating objects in headers and footers are not exported to PDF yet.',
			);
			continue;
		}
		if (!run.field) {
			runs.push(run);
			continue;
		}
		const { field: _field, ...plain } = run;
		runs.push({ ...plain, text: fieldDisplayText(run, values) });
	}
	return { ...paragraph, runs };
}

function fillBlocks(
	blocks: readonly Block[],
	values: PageFieldContext,
	mode: ReviewDisplayMode,
	report: DocxPdfReportBuilder,
): Block[] {
	return blocks.map((block) =>
		block.type === 'paragraph'
			? fillParagraph(block, values, mode, report)
			: {
					...block,
					rows: block.rows.map((row) =>
						row.map((cell) => ({
							...cell,
							paragraphs: cell.paragraphs.map((p) => fillParagraph(p, values, mode, report)),
						})),
					),
				},
	);
}

/** Lays out every page's header and footer. */
export function placeStories(
	model: DocumentModel,
	pages: readonly LayoutPageBox[],
	measurer: TextMeasurer,
	mode: ReviewDisplayMode,
	report: DocxPdfReportBuilder,
	now: Date,
): PlacedStory[][] {
	const labels = pageNumbers(model, [...pages]);
	const numbers = pageNumberValues(model, [...pages]);
	const counts = sectionPageCounts([...pages]);
	const styleRefs = pageStyleRefs(model, [...pages]);
	const cache = new Map<string, { blocks: LayoutBlockBox[]; height: number }>();
	const layout = (blocks: Block[], widthPx: number, sectionIndex: number) => {
		const key = `${widthPx}\u0000${sectionIndex}\u0000${JSON.stringify(blocks)}`;
		const hit = cache.get(key);
		if (hit) return hit;
		const section = model.sections?.[sectionIndex];
		const story: DocumentModel = { ...model, blocks, footnotes: [], endnotes: [] };
		if (section) {
			// The story is laid out as a one-page document as wide as the text column, with none of
			// the section's own decorations.
			const {
				headers: _headers,
				footers: _footers,
				lineNumbering: _lineNumbering,
				lineNumberSettings: _lineNumberSettings,
				pageBorders: _pageBorders,
				verticalAlign: _verticalAlign,
				...plain
			} = section;
			story.sections = [
				{
					...plain,
					endsAtBlockId: blocks.at(-1)?.id ?? '',
					type: 'nextPage',
					pageWidthTwips: twips(Math.max(1, Math.round(widthPx * PX_TWIPS))),
					pageHeightTwips: twips(TALL_PAGE_TWIPS),
					marginTopTwips: signedTwips(0),
					marginBottomTwips: signedTwips(0),
					marginLeftTwips: twips(0),
					marginRightTwips: twips(0),
					columns: { count: 1, equalWidth: true },
				},
			];
		} else {
			delete story.sections;
			story.page = {
				width: widthPx,
				height: TALL_PAGE_TWIPS / PX_TWIPS,
				marginTop: 0,
				marginRight: 0,
				marginBottom: 0,
				marginLeft: 0,
			};
		}
		const result = layoutDocumentModel(story, measurer, { reviewDisplayMode: mode });
		const placed = result.pages[0]?.columns[0]?.blocks ?? [];
		const height = placed.reduce(
			(bottom, block) => Math.max(bottom, block.yPx + block.heightPx),
			0,
		);
		const entry = { blocks: placed, height };
		cache.set(key, entry);
		return entry;
	};
	return pages.map((page, index) => {
		const section = model.sections?.[page.sectionIndex];
		const values: PageFieldContext = {
			page: labels[index] ?? String(index + 1),
			numPages: String(pages.length),
			sectionPages: String(counts.get(page.sectionIndex) ?? 1),
			now,
		};
		const widthPx = page.widthPx - page.marginLeftPx - page.marginRightPx;
		const placed: PlacedStory[] = [];
		for (const kind of ['header', 'footer'] as const) {
			const content: HeaderFooterContent | undefined = headerFooterForPage(
				model,
				page,
				numbers[index] ?? index + 1,
				`${kind}s`,
			);
			if (!content?.blocks.length) continue;
			const blocks = fillBlocks(styleRefs.content(content, index).blocks, values, mode, report);
			const { blocks: boxes, height } = layout(blocks, widthPx, page.sectionIndex);
			const distanceTwips =
				kind === 'header' ? section?.headerDistanceTwips : section?.footerDistanceTwips;
			const distancePx = (distanceTwips ?? 720) / PX_TWIPS;
			placed.push({
				kind,
				x: page.marginLeftPx,
				y: kind === 'header' ? distancePx : page.heightPx - distancePx - height,
				blocks: boxes,
			});
		}
		return placed;
	});
}
