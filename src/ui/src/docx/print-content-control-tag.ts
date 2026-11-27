/**
 * Draws a content control tag in Print Layout as the outline Word prints: an inline SVG in
 * points (its view box) holding the arc-sided path, the paragraph, row or cell marker and the
 * title on Word's baseline. The span keeps its laid-out box; the title stays its text.
 */
import {
	contentControlTagMarker,
	contentControlTagMetrics,
	contentControlTagOutline,
	contentControlTagOutlineWidth,
	contentControlTagStyle,
	contentControlTagTitleOffset,
	tagPathData,
	type LayoutFragment,
} from 'ooxml-core/docx/layout';

const SVG = 'http://www.w3.org/2000/svg';
const num = (value: number) => String(Number(value.toFixed(3)));

function svgElement(doc: Document, name: string, attributes: Record<string, string>): Element {
	const element = doc.createElementNS(SVG, name);
	for (const [key, value] of Object.entries(attributes)) element.setAttribute(key, value);
	return element;
}

/** Replaces `span`'s text with the outline and title of `fragment`'s tag. */
export function paintPrintContentControlTag(span: HTMLElement, fragment: LayoutFragment): void {
	const edge = fragment.contentControlTag;
	if (!edge) return;
	const doc = span.ownerDocument;
	const structural = !!fragment.blockContentControlTag;
	const style = contentControlTagStyle;
	const advancePt = (fragment.widthPx * 72) / 96;
	const heightPt = contentControlTagMetrics.boxHeightPt;
	const widthPt = contentControlTagOutlineWidth(edge, advancePt, structural);
	const svg = svgElement(doc, 'svg', {
		class: 'dve-print-content-control-tag-shape',
		viewBox: `0 0 ${num(advancePt)} ${num(heightPt)}`,
		width: num(fragment.widthPx),
		height: num((heightPt * 96) / 72),
		overflow: 'visible',
	});
	svg.append(
		svgElement(doc, 'path', {
			d: tagPathData(contentControlTagOutline(edge, widthPt)),
			fill: structural ? style.blockFill : style.fill,
			'fill-rule': 'evenodd',
			stroke: style.stroke,
			'stroke-width': num(style.strokeWidthPt),
			'stroke-linejoin': 'round',
		}),
	);
	if (structural) {
		const { arrow, bar } = contentControlTagMarker(edge, widthPt);
		const marker = {
			fill: style.stroke,
			stroke: style.stroke,
			'stroke-width': num(style.marker.strokeWidthPt),
		};
		svg.append(
			svgElement(doc, 'path', {
				d: `M ${arrow.map(([x, y]) => `${num(x)} ${num(y)}`).join(' L ')} Z`,
				...marker,
			}),
			svgElement(doc, 'rect', {
				x: num(bar.x),
				y: num(bar.y),
				width: num(bar.width),
				height: num(bar.height),
				...marker,
			}),
		);
	}
	const title = svgElement(doc, 'text', {
		x: num(contentControlTagTitleOffset(edge, structural)),
		y: num(style.titleBaselinePt),
		fill: style.text,
		'font-size': num(contentControlTagMetrics.fontSizePt),
	});
	title.textContent = fragment.text;
	svg.append(title);
	span.replaceChildren(svg);
}
