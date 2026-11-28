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

/** Word draws compound borders in steps of 0.12 pt (1/600 in), in px. */
const UNIT_PX = (0.12 * 96) / 72;

/**
 * Band and gap widths of the compound line styles, from the top (left) edge of the border to its
 * inner side, alternating band and gap, in px. The same order is drawn for all four sides: Word
 * does not mirror a bottom or right border. Fitted to Word's own PDF export at `w:sz` 1 to 48
 * (#388, `issue-388-word-compound-bands.json`):
 *
 * - `w` = floor(sz x 25 / 24) steps of 0.12 pt (the nominal sz / 8 pt, floored to the step).
 * - `double` and `triple`: bands and gaps of `w`.
 * - small gap: a 0.72 pt thin band and gap, one `w` band.
 * - medium gap: the thin band and the gap are half of `w` (floor(sz x 25 / 48) steps).
 * - large gap: a 1.44 pt band, a 0.72 pt thin band, and a gap of `w`.
 * - the thin-thick styles put the `w`-wide (or 1.44 pt) band first, the thick-thin styles last.
 */
function compoundBorder(style: string | undefined, sizeEighthPoints: number): number[] | undefined {
	const steps = (count: number) => Math.max(1, Math.floor(count)) * UNIT_PX;
	const w = steps((sizeEighthPoints * 25) / 24);
	if (style === 'triple') return [w, w, w, w, w];
	if (style === 'double') return [w, w, w];
	const match = /^(thinThickThin|thinThick|thickThin)(Small|Medium|Large)Gap$/.exec(style ?? '');
	if (!match) return undefined;
	const small = steps(6);
	const half = steps((sizeEighthPoints * 25) / 48);
	const [big, thin, gap] =
		match[2] === 'Small'
			? [w, small, small]
			: match[2] === 'Medium'
				? [w, half, half]
				: [steps(12), small, w];
	// Word's "thin thick" lists the wide band first; "thick thin" lists it last.
	if (match[1] === 'thinThickThin') {
		const outer = match[2] === 'Medium' ? half : thin;
		return [outer, gap, match[2] === 'Large' ? steps(12) : big, gap, outer];
	}
	return match[1] === 'thinThick' ? [big, gap, thin] : [thin, gap, big];
}

/** Bands of a compound border, measured from its top (left) edge, in the order Word draws them. */
export function paragraphBorderBands(line: LayoutParagraphBorder) {
	let offsetPx = 0;
	const bands: { offsetPx: number; widthPx: number }[] = [];
	line.compound?.forEach((widthPx, index) => {
		if (index % 2 === 0) bands.push({ offsetPx, widthPx });
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
		const compound = compoundBorder(source?.style, source?.sizeEighthPoints ?? 4);
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
