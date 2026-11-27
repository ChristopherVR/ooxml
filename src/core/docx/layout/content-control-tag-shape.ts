/**
 * The outline Word draws for a content control tag, read from the vector paths of its
 * print-optimized PDF export (`fixtures/word-layout-reports/README.md`). Every length is in
 * points in a y-down space whose origin is the top left of the tag box (`boxHeightPt` high).
 * DOM-free: the PDF painter, Print Layout and the editor all draw these numbers.
 */
import { contentControlTagMetrics } from './content-control-tag-metrics';

export const contentControlTagStyle = {
	/** Inline tag fill, PDF `0.925 0.949 0.969 rg`. */
	fill: '#ecf2f7',
	/** Paragraph, row and cell tag fill, PDF `0.988 g`. */
	blockFill: '#fcfcfc',
	/** Outline and marker colour, PDF `0.627 0.745 0.843 RG`. */
	stroke: '#a0bed7',
	/** Title colour, PDF `0 g`. */
	text: '#000000',
	strokeWidthPt: 1.08,
	/** Both side edges are arcs of one circle bowing this far across the tag. */
	edgeDepthPt: 2.28,
	/** An unlabelled inline closing tag's outline width; the rest of its advance is a gap. */
	inlineEndOutlinePt: 9.84,
	/** Where the title starts, from the tag's left edge. */
	titleOffsetPt: { inline: 6.72, blockStart: 9, end: 5.04 },
	/** The title baseline, from the top of the box (3 pt above the body baseline). */
	titleBaselinePt: contentControlTagMetrics.boxAscentPt - contentControlTagMetrics.labelRisePt,
	/** Tahoma's ascent and descent per em, from the font descriptor of Word's PDF. */
	fontAscentEm: 1,
	fontDescentEm: 0.207,
	/** The arrow and bar marking a paragraph, row or cell tag. */
	marker: {
		strokeWidthPt: 0.14,
		topPt: 5.1,
		heightPt: 3.36,
		barWidthPt: 0.6,
		/** From the start tag's left edge: arrow tip, arrow base and bar. */
		start: { tipPt: 2.34, basePt: 4.5, barPt: 5.58 },
		/** From the end tag's right edge: arrow tip, arrow base and bar's left side. */
		end: { tipPt: 2.82, basePt: 4.98, barPt: 6.66 },
	},
} as const;

export type TagPathCommand =
	| readonly ['M' | 'L', number, number]
	| readonly ['C', number, number, number, number, number, number];

export interface TagMarker {
	/** Arrow tip, then the two ends of its base. */
	arrow: readonly (readonly [number, number])[];
	bar: { x: number; y: number; width: number; height: number };
}

/** One edge: a circular arc from `(x, fromY)` to `(x, toY)` bowing `side` (-1 left, 1 right). */
export function tagEdgeArc(x: number, fromY: number, toY: number, side: -1 | 1): TagPathCommand {
	const depth = contentControlTagStyle.edgeDepthPt;
	const half = Math.abs(toY - fromY) / 2;
	const radius = (half * half + depth * depth) / (2 * depth);
	const sweep = 2 * Math.asin(half / radius);
	const handle = (4 / 3) * Math.tan(sweep / 4) * radius;
	const across = (side * handle * half) / radius;
	const along = (Math.sign(toY - fromY) * handle * (radius - depth)) / radius;
	return ['C', x + across, fromY + along, x + across, toY - along, x, toY];
}

/** The width of a tag's outline: a labelled tag fills its advance, an inline end does not. */
export function contentControlTagOutlineWidth(
	edge: 'start' | 'end',
	advancePt: number,
	structural: boolean,
): number {
	return edge === 'end' && !structural
		? Math.min(advancePt, contentControlTagStyle.inlineEndOutlinePt)
		: advancePt;
}

/**
 * The closed outline of a tag `widthPt` wide (see `contentControlTagOutlineWidth`). An opening
 * tag's edges both bow to the left, a closing tag's both to the right, so the pair reads as
 * brackets around the content.
 */
export function contentControlTagOutline(edge: 'start' | 'end', widthPt: number): TagPathCommand[] {
	const height = contentControlTagMetrics.boxHeightPt;
	const depth = Math.min(contentControlTagStyle.edgeDepthPt, widthPt / 2);
	const width = Math.max(widthPt, 2 * depth);
	if (edge === 'start')
		return [
			['M', depth, height],
			tagEdgeArc(depth, height, 0, -1),
			['L', width, 0],
			tagEdgeArc(width, 0, height, -1),
		];
	return [
		['M', 0, height],
		tagEdgeArc(0, height, 0, 1),
		['L', width - depth, 0],
		tagEdgeArc(width - depth, 0, height, 1),
	];
}

/** Where the title of a tag starts, from its left edge. */
export function contentControlTagTitleOffset(edge: 'start' | 'end', structural: boolean): number {
	const { titleOffsetPt } = contentControlTagStyle;
	if (edge === 'end') return titleOffsetPt.end;
	return structural ? titleOffsetPt.blockStart : titleOffsetPt.inline;
}

/** The arrow and bar a paragraph, row or cell tag `widthPt` wide carries. */
export function contentControlTagMarker(edge: 'start' | 'end', widthPt: number): TagMarker {
	const { marker } = contentControlTagStyle;
	const top = marker.topPt;
	const bottom = top + marker.heightPt;
	const middle = top + marker.heightPt / 2;
	const from = edge === 'start' ? (x: number) => x : (x: number) => widthPt - x;
	const place = edge === 'start' ? marker.start : marker.end;
	return {
		arrow: [
			[from(place.tipPt), middle],
			[from(place.basePt), top],
			[from(place.basePt), bottom],
		],
		bar: { x: from(place.barPt), y: top, width: marker.barWidthPt, height: marker.heightPt },
	};
}

const round = (value: number) => Number(value.toFixed(3));

/** SVG path data for an outline (closed unless `closed` is false), scaled from points. */
export function tagPathData(
	commands: readonly TagPathCommand[],
	{ scale = 1, closed = true }: { scale?: number; closed?: boolean } = {},
): string {
	const parts = commands.map(([op, ...values]) =>
		[op, ...values.map((value) => round(value * scale))].join(' '),
	);
	return closed ? `${parts.join(' ')} Z` : parts.join(' ');
}
