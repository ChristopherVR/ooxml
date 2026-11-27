/**
 * The editing view's content control tag shape. A tag's width follows its title there, so the
 * outline is split: the two arc-sided ends are small SVGs of fixed size at each side of the tag,
 * and the straight middle is the tag's own CSS background (`styles/content-controls.css`).
 * Their colours are CSS custom properties so themes can restyle them.
 */
import {
	contentControlTagMarker,
	contentControlTagMetrics,
	contentControlTagStyle,
	contentControlTagTitleOffset,
	tagEdgeArc,
	tagPathData,
	type TagPathCommand,
} from 'ooxml-core/docx/layout';
import type { TagEdge } from './content-control-tag-widgets';

const SVG = 'http://www.w3.org/2000/svg';
const num = (value: number) => String(Number(value.toFixed(3)));
const pt = (value: number) => `${num(value)}pt`;

function svgElement(doc: Document, name: string, attributes: Record<string, string>): Element {
	const element = doc.createElementNS(SVG, name);
	for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
	return element;
}

/** The fill region and stroked line of one end, in its own `edgeDepthPt` wide box. */
function end(edge: TagEdge, side: 'left' | 'right'): { fill: string; line: string } {
	const d = contentControlTagStyle.edgeDepthPt;
	const h = contentControlTagMetrics.boxHeightPt;
	const bow = edge === 'start' ? -1 : 1;
	// Opening tags draw their arcs at the right of each end, closing tags at the left.
	const x = edge === 'start' ? d : 0;
	const down = tagEdgeArc(x, 0, h, bow);
	const up = tagEdgeArc(x, h, 0, bow);
	const path = (commands: TagPathCommand[], closed = false) => tagPathData(commands, { closed });
	// The ends of a tag's top and bottom lines that the middle does not cover.
	const stubbed = (edge === 'start') === (side === 'right');
	if (!stubbed) return { fill: path([['M', x, h], up], true), line: path([['M', x, 0], down]) };
	const other = d - x;
	return {
		fill: path([['M', other, 0], ['L', x, 0], down, ['L', other, h]], true),
		line: path([['M', other, 0], ['L', x, 0], down, ['L', other, h]]),
	};
}

function marker(doc: Document, edge: TagEdge, side: 'left' | 'right'): Element[] {
	const d = contentControlTagStyle.edgeDepthPt;
	// Placed in an end's box: the opening tag's left end, the closing tag's right one.
	const { arrow, bar } = contentControlTagMarker(edge, d);
	if ((edge === 'start') !== (side === 'left')) return [];
	return [
		svgElement(doc, 'path', {
			class: 'dve-sdt-tag-marker',
			d: `M ${arrow.map(([x, y]) => `${num(x)} ${num(y)}`).join(' L ')} Z`,
		}),
		svgElement(doc, 'rect', {
			class: 'dve-sdt-tag-marker',
			x: num(bar.x),
			y: num(bar.y),
			width: num(bar.width),
			height: num(bar.height),
		}),
	];
}

/** The two ends of a tag, with the marker of a paragraph, row or cell tag. */
export function tagEnds(doc: Document, edge: TagEdge, inline: boolean): Element[] {
	const { edgeDepthPt: d, strokeWidthPt } = contentControlTagStyle;
	const h = contentControlTagMetrics.boxHeightPt;
	return (['left', 'right'] as const).map((side) => {
		const svg = svgElement(doc, 'svg', {
			class: 'dve-sdt-tag-end',
			'data-side': side,
			viewBox: `0 ${num(-strokeWidthPt / 2)} ${num(d)} ${num(h + strokeWidthPt)}`,
			overflow: 'visible',
			'aria-hidden': 'true',
		});
		const { fill, line } = end(edge, side);
		svg.append(
			svgElement(doc, 'path', { class: 'dve-sdt-tag-fill', d: fill }),
			svgElement(doc, 'path', { class: 'dve-sdt-tag-line', d: line }),
			...(inline ? [] : marker(doc, edge, side)),
		);
		return svg;
	});
}

/**
 * The tag box's own style: Word's advance (title offset, padding and minimum), a box that puts
 * the outline 1.68 pt below the body baseline and the title 3 pt above it, and margins that give
 * a tagged line the same minimum height as in Print Layout, so lines do not jump between views.
 */
export function tagBoxStyle(edge: TagEdge, inline: boolean): Partial<CSSStyleDeclaration> {
	const metrics = contentControlTagMetrics;
	const style = contentControlTagStyle;
	const labelled = edge === 'start' || !inline;
	const offset = contentControlTagTitleOffset(edge, !inline);
	// The box includes the stroke's outer half above and below the outline.
	const boxHeight = metrics.boxHeightPt + style.strokeWidthPt;
	const baseline = style.titleBaselinePt + style.strokeWidthPt / 2;
	// A single line box puts the baseline half the line height plus half (ascent - descent) down.
	const lineHeight = 2 * baseline - (style.fontAscentEm - style.fontDescentEm) * metrics.fontSizePt;
	// Margins stretch the line to Print Layout's tagged ascent and descent (line-metrics.ts).
	const above = metrics.labelRisePt + baseline;
	const below = boxHeight - baseline - metrics.labelRisePt;
	return {
		height: pt(boxHeight),
		lineHeight: pt(lineHeight),
		verticalAlign: pt(metrics.labelRisePt),
		marginTop: pt(metrics.lineAscentPt - above),
		marginBottom: pt(metrics.lineDescentPt - below),
		...(labelled
			? {
					minInlineSize: pt(metrics.minLabelWidthPt),
					paddingInline: `${pt(offset)} ${pt(metrics.labelPaddingPt - offset)}`,
				}
			: {
					minInlineSize: pt(style.inlineEndOutlinePt),
					marginInlineEnd: pt(metrics.inlineEndWidthPt - style.inlineEndOutlinePt),
				}),
	};
}
