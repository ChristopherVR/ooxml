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

/** Stroke/gap widths for the compound line styles, in multiples of the nominal stroke. */
function compoundBorder(style: string | undefined, stroke: number): number[] | undefined {
	if (style === 'triple') return [stroke, stroke, stroke, stroke, stroke];
	const match = /^(thinThickThin|thinThick|thickThin)(Small|Medium|Large)Gap$/.exec(style ?? '');
	if (!match) return undefined;
	const thick = stroke * 1.5;
	const gap = stroke * (match[2] === 'Small' ? 1.5 : match[2] === 'Medium' ? 2 : 3);
	return match[1] === 'thinThickThin'
		? [stroke, gap, thick, gap, stroke]
		: match[1] === 'thinThick'
			? [stroke, gap, thick]
			: [thick, gap, stroke];
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
		const compound = compoundBorder(source?.style, line.widthPx);
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
	// Keys come from the original borders, before any paragraph in the group is adjusted.
	const keys = blocks.map((block) =>
		block.kind === 'paragraph' && block.borders ? JSON.stringify(block.borders) : undefined,
	);
	blocks.forEach((block, index) => {
		if (block.kind !== 'paragraph' || !block.borders) return;
		const sameAsPrevious = keys[index - 1] === keys[index];
		const sameAsNext = keys[index + 1] === keys[index];
		if (!sameAsPrevious && !sameAsNext) return;
		const borders = { ...block.borders };
		if (sameAsNext) {
			delete borders.bottom;
			block.borderJoinNext = true;
		}
		if (sameAsPrevious) {
			const line = between.get(block);
			if (line) borders.top = line;
			else delete borders.top;
		}
		block.borders = borders;
	});
}
