/**
 * What a page's header and footer draw outside their text flow, for the Word PDF export: floating
 * pictures, shapes, charts and text boxes, placed on the page from their anchor paragraph as Print
 * Layout places them (`print-header-footer.ts`), and the header's watermark. DOM-free.
 */
import {
	isRunHiddenForReview,
	type DocumentModel,
	type HeaderFooterContent,
	type Paragraph,
	type PictureWatermarkSpec,
	type ReviewDisplayMode,
	type WatermarkSpec,
} from 'ooxml-core/docx';
import {
	floatPosition,
	paragraphFloats,
	storyTextBoxLayouts,
	type LayoutPageBox,
	type TextBoxLayouts,
	type TextMeasurer,
} from 'ooxml-core/docx/layout';
import type { DocxPdfPainter } from './pdf-blocks';
import type { DocxPdfPicture } from './pdf-pictures';
import type { DocxPdfReportBuilder } from './pdf-report';
import type { PlacedStory } from './pdf-stories';

/** A header or footer object, placed on the page. */
export interface DocxPdfStoryFloat extends DocxPdfPicture {
	kind: 'header' | 'footer';
	xPx: number;
	yPx: number;
	behindText: boolean;
}

/** A page's header and footer graphics. */
export interface DocxPdfStoryGraphics {
	floats: DocxPdfStoryFloat[];
	watermark?: WatermarkSpec;
	pictureWatermark?: PictureWatermarkSpec;
}

function storyParagraphs(content: HeaderFooterContent): Paragraph[] {
	return content.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
}

/** Every page's header and footer graphics. */
export function storyGraphics(
	model: DocumentModel,
	pages: readonly LayoutPageBox[],
	stories: readonly (readonly PlacedStory[])[],
	measurer: TextMeasurer,
	mode: ReviewDisplayMode,
	report: DocxPdfReportBuilder,
): DocxPdfStoryGraphics[] {
	const hidden = (run: Paragraph['runs'][number]) => isRunHiddenForReview(run, mode);
	const textBoxes = new Map<HeaderFooterContent, TextBoxLayouts>();
	const textBoxesOf = (content: HeaderFooterContent) => {
		let layouts = textBoxes.get(content);
		if (!layouts) {
			layouts = storyTextBoxLayouts(
				model,
				content.blocks,
				measurer,
				{ reviewDisplayMode: mode },
				(m) => report.note(m),
			);
			textBoxes.set(content, layouts);
		}
		return layouts;
	};
	return pages.map((page, index) => {
		const graphics: DocxPdfStoryGraphics = { floats: [] };
		const column = { xPx: 0, widthPx: page.widthPx - page.marginLeftPx - page.marginRightPx };
		for (const story of stories[index] ?? []) {
			const paragraphs = storyParagraphs(story.content);
			if (story.kind === 'header') {
				const image = paragraphs
					.flatMap((paragraph) => paragraph.runs)
					.find(
						(run) => !hidden(run) && (run.image?.watermark || run.image?.pictureWatermark),
					)?.image;
				if (image?.pictureWatermark) graphics.pictureWatermark = image.pictureWatermark;
				else if (image?.watermark?.text) graphics.watermark = image.watermark;
			}
			for (const paragraph of paragraphs) {
				if (!paragraph.runs.some((run) => run.image?.anchored)) continue;
				// The anchor paragraph where the story laid it out; a paragraph of a table cell
				// anchors at the story's top.
				const box = story.blocks.find((block) => block.blockId === paragraph.id);
				const anchor = {
					topPx: story.y + (box?.yPx ?? 0),
					heightPx: box?.heightPx ?? 0,
				};
				const floats = paragraphFloats(
					paragraph,
					model.theme,
					textBoxesOf(story.content),
					false,
					hidden,
				);
				for (const float of floats) {
					const { xPx, yPx } = floatPosition(float, page, column, anchor);
					graphics.floats.push({
						...float,
						kind: story.kind,
						xPx,
						yPx,
						behindText: !!float.behindText,
					});
				}
			}
		}
		return graphics;
	});
}

/** The pictures a page's header and footer graphics draw, for loading before painting. */
export function* storyGraphicPictures(graphics: DocxPdfStoryGraphics): Generator<DocxPdfPicture> {
	yield* graphics.floats;
	const picture = graphics.pictureWatermark;
	if (picture)
		yield {
			partName: picture.partName,
			contentType: picture.contentType,
			widthPx: (picture.widthPt * 4) / 3,
			heightPx: (picture.heightPt * 4) / 3,
		};
}

/** Draws the floating objects of one story, those behind the text or those in front of it. */
export function paintStoryFloats(
	painter: DocxPdfPainter,
	graphics: DocxPdfStoryGraphics | undefined,
	kind: 'header' | 'footer',
	behindText: boolean,
): void {
	for (const float of graphics?.floats ?? [])
		if (float.kind === kind && float.behindText === behindText)
			painter.pictures.draw(painter.page, float, float.xPx, float.yPx);
}
