/**
 * Which slides a vector PDF holds and how big its pages are, like PowerPoint's Save as PDF:
 * hidden slides are left out unless asked for, a range counts every slide (hidden ones
 * included) and each page is the slide's own size, not a paper size. Pure.
 */
import type { CanvasSize } from '../types';
import type { SlideSizeEmu } from '../render/slide-size';

/** Which slides to export. */
export interface PdfSlideSelection {
	/** Include slides marked hidden. Default `false`, as PowerPoint's "Include hidden slides". */
	includeHiddenSlides?: boolean;
	/** First slide, 1-based, inclusive. Default 1. */
	from?: number;
	/** Last slide, 1-based, inclusive. Default the last slide. */
	to?: number;
	/** Explicit 0-based slide indexes (current slide, a selection); overrides `from`/`to`. */
	slideIndexes?: readonly number[];
}

/** The most pages one export writes. */
export const MAX_PDF_SLIDES = 2000;

/** The 0-based indexes of the slides to export, in deck order. */
export function selectPdfSlides(
	slides: readonly { hidden?: boolean }[],
	selection: PdfSlideSelection = {},
): number[] {
	const count = slides.length;
	let indexes: number[];
	if (selection.slideIndexes) {
		indexes = [...new Set(selection.slideIndexes)]
			.filter((index) => Number.isInteger(index) && index >= 0 && index < count)
			.sort((a, b) => a - b);
	} else {
		const from = Math.max(1, Math.floor(selection.from ?? 1));
		const to = Math.min(count, Math.floor(selection.to ?? count));
		indexes = [];
		for (let index = from - 1; index < to; index++) indexes.push(index);
	}
	if (!selection.includeHiddenSlides) indexes = indexes.filter((index) => !slides[index]?.hidden);
	return indexes.slice(0, MAX_PDF_SLIDES);
}

/** EMU per PDF point (12700), and CSS px (at 96 per inch) per point. */
const EMU_PER_POINT = 12700;
const PX_PER_POINT = 96 / 72;

/** The PDF page size in points: the slide size in EMU when known, else the canvas size. */
export function slidePdfPageSize(
	canvasSize: CanvasSize,
	slideSize?: Pick<SlideSizeEmu, 'widthEmu' | 'heightEmu'>,
): { width: number; height: number } {
	if (slideSize && slideSize.widthEmu > 0 && slideSize.heightEmu > 0)
		return {
			width: slideSize.widthEmu / EMU_PER_POINT,
			height: slideSize.heightEmu / EMU_PER_POINT,
		};
	return { width: canvasSize.width / PX_PER_POINT, height: canvasSize.height / PX_PER_POINT };
}
