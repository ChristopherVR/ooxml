/**
 * What a page's header and footer draw outside their text flow, for the Word PDF export: floating
 * pictures, shapes, charts and text boxes, placed on the page from their anchor paragraph as Print
 * Layout places them (`print-header-footer.ts`), and the header's watermark. DOM-free.
 */
import {
	isRunHiddenForReview,
	type HeaderFooterContent,
	type Paragraph,
	type PictureWatermarkSpec,
	type ReviewDisplayMode,
	type WatermarkSpec,
} from 'ooxml-core/docx';
import type { PlacedStory } from 'ooxml-core/docx/layout';
import type { DocxPdfPainter } from './pdf-blocks';
import type { DocxPdfPicture } from './pdf-pictures';

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
	stories: readonly (readonly PlacedStory[])[],
	mode: ReviewDisplayMode,
): DocxPdfStoryGraphics[] {
	return stories.map((pageStories) => {
		const graphics: DocxPdfStoryGraphics = { floats: [] };
		for (const story of pageStories) {
			const paragraphs = storyParagraphs(story.content);
			if (story.kind === 'header') {
				const image = paragraphs
					.flatMap((paragraph) => paragraph.runs)
					.find(
						(run) =>
							!isRunHiddenForReview(run, mode) &&
							(run.image?.watermark || run.image?.pictureWatermark),
					)?.image;
				if (image?.pictureWatermark) graphics.pictureWatermark = image.pictureWatermark;
				else if (image?.watermark?.text) graphics.watermark = image.watermark;
			}
			for (const float of story.floats)
				graphics.floats.push({ ...float, kind: story.kind, behindText: !!float.behindText });
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
