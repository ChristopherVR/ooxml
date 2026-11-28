import type { DocumentModel, ParagraphFormatting } from '../index';
import { percentageShadingColor } from '../percentage-shading';
import { border, cssHex } from './adapt-table';
import type {
	LayoutBlock,
	LayoutParagraph,
	LayoutParagraphBorder,
	LayoutParagraphBorders,
} from './input';

const pointsToPx = (points: number) => (points * 96) / 72;

/** Word's thin band and its small gap: 0.72 pt, whatever the border's `w:sz` (#388). */
const THIN_POINTS = 0.72;

/**
 * Band and gap widths of the compound line styles, outermost first, alternating band and gap.
 * In Word `w:sz` is the width of the thick band of the thin-thick styles: the thin band and the
 * small gap stay at 0.72 pt (a `w:sz` 24 `thinThickSmallGap` is 3.00 + 0.72 + 0.72 pt). A `double`
 * draws two bands of `w:sz` around a gap of `w:sz`. The medium gap widens with `w:sz` (measured
 * at `w:sz` 12 and 24) and the large gap is 1.44 pt.
 */
function compoundBorder(style: string | undefined, stroke: number): number[] | undefined {
	if (style === 'triple') return [stroke, stroke, stroke, stroke, stroke];
	if (style === 'double') return [stroke, stroke, stroke];
	const match = /^(thinThickThin|thinThick|thickThin)(Small|Medium|Large)Gap$/.exec(style ?? '');
	if (!match) return undefined;
	const thin = pointsToPx(THIN_POINTS);
	const gap =
		match[2] === 'Small'
			? thin
			: match[2] === 'Medium'
				? Math.max(thin, stroke * 0.96 - thin)
				: thin * 2;
	return match[1] === 'thinThickThin'
		? [thin, gap, stroke, gap, thin]
		: match[1] === 'thinThick'
			? [thin, gap, stroke]
			: [stroke, gap, thin];
}

/** Visible bands measured from the outside or the inside edge of a compound border. */
export function paragraphBorderBands(line: LayoutParagraphBorder, reverse: boolean) {
	let offsetPx = 0;
	const bands: { offsetPx: number; widthPx: number }[] = [];
	line.compound?.forEach((widthPx, index) => {
		if (index % 2 === 0)
			bands.push({ offsetPx: reverse ? line.widthPx - offsetPx - widthPx : offsetPx, widthPx });
		offsetPx += widthPx;
	});
	return bands;
}

/** Paragraph borders and shading from resolved formatting. */
export function paragraphBox(
	formatting: ParagraphFormatting,
	theme: DocumentModel['theme'],
): Pick<LayoutParagraph, 'borders' | 'shading'> & {
	betweenBorder?: LayoutParagraphBorders['top'];
} {
	const borders: LayoutParagraphBorders = {};
	let betweenBorder: LayoutParagraphBorders['top'];
	for (const side of ['top', 'bottom', 'left', 'right', 'between'] as const) {
		const source = formatting.borders?.[side];
		const line = border(source, theme, true);
		if (!line) continue;
		const compound = compoundBorder(source?.style, pointsToPx((source?.sizeEighthPoints ?? 4) / 8));
		const resolved = {
			...line,
			spacePx: pointsToPx(source?.spacePoints ?? 0),
			...(source?.shadow ? { shadow: true as const } : {}),
			...(compound ? { compound, widthPx: compound.reduce((sum, width) => sum + width, 0) } : {}),
		};
		if (side === 'between') betweenBorder = resolved;
		else borders[side] = resolved;
	}
	const shading = cssHex(
		percentageShadingColor(formatting.shadingFill, formatting.percentageShading),
	);
	return {
		...(Object.keys(borders).length ? { borders } : {}),
		...(shading ? { shading } : {}),
		...(betweenBorder ? { betweenBorder } : {}),
	};
}

/**
 * Word draws consecutive paragraphs with identical borders as one box: the top line only above
 * the first, the bottom line only below the last, and the `between` line (if any) between them.
 */
export function groupParagraphBorders(
	blocks: LayoutBlock[],
	between: Map<LayoutParagraph, LayoutParagraphBorders['top']>,
): void {
	// Keys come from the original borders, before any paragraph in the group is adjusted. A
	// paragraph with only a `between` line is grouped too (#417).
	const keys = blocks.map((block) => {
		if (block.kind !== 'paragraph') return undefined;
		const line = between.get(block);
		return block.borders || line ? JSON.stringify([block.borders ?? {}, line ?? null]) : undefined;
	});
	blocks.forEach((block, index) => {
		if (block.kind !== 'paragraph' || keys[index] === undefined) return;
		const sameAsPrevious = keys[index - 1] === keys[index];
		const sameAsNext = keys[index + 1] === keys[index];
		if (!sameAsPrevious && !sameAsNext) return;
		const borders = { ...block.borders };
		const line = between.get(block);
		if (sameAsNext) {
			delete borders.bottom;
			block.borderJoinNext = true;
			// Word leaves the `w:space` of the between line under this paragraph; the line and its
			// second gap open the next one, so only the first gap counts at a page's end.
			if (line) block.borderGapBelowPx = line.spacePx;
		}
		if (sameAsPrevious) {
			if (line) borders.top = line;
			else delete borders.top;
		}
		if (Object.keys(borders).length) block.borders = borders;
		else delete block.borders;
	});
}
