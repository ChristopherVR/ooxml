import type { PptxElement, ShapeStyle } from 'ooxml-core/pptx';
import { hasShapeProperties } from 'ooxml-core/pptx';

import { markerPath, normalizeArrow } from './connector-markers';
import { escapeSvgAttr } from './visual-effects';
import { hasStrokePaint } from './stroke-paint';

/** Open, unfilled freeforms need an unclipped SVG stroke, including their ends. */
export function isArrowedCustomPath(element: PptxElement): boolean {
	if (element.type !== 'shape' || !element.pathData || !element.pathWidth || !element.pathHeight) {
		return false;
	}
	const style = element.shapeStyle;
	if (!hasStrokePaint(style) || (style?.strokeWidth ?? 0) <= 0) return false;
	return (
		!/[zZ]/u.test(element.pathData) &&
		(style?.fillMode === 'none' || style?.fillColor === 'transparent') &&
		Boolean(normalizeArrow(style?.connectorStartArrow) || normalizeArrow(style?.connectorEndArrow))
	);
}

export interface CustomPathMarkers {
	markup: string;
	start: string | undefined;
	end: string | undefined;
}

/** Generated SVG only; all document-derived attribute values are escaped. */
export function buildCustomPathMarkers(
	element: PptxElement,
	stroke: string,
): CustomPathMarkers | undefined {
	if (!isArrowedCustomPath(element) || !hasShapeProperties(element)) return undefined;
	const style: ShapeStyle = element.shapeStyle ?? {};
	// Encode each code point so distinct shape IDs cannot collapse to the same SVG ID.
	const seed = [...element.id].map((char) => char.codePointAt(0)!.toString(16)).join('-');
	const markup: string[] = [];
	const add = (end: 'Start' | 'End'): string | undefined => {
		const arrow = normalizeArrow(style[`connector${end}Arrow`]);
		if (!arrow) return undefined;
		const shape = markerPath(
			arrow,
			style[`connector${end}ArrowWidth`],
			style[`connector${end}ArrowLength`],
		);
		const id = `freeform-${seed}-${end}`;
		const color = escapeSvgAttr(stroke);
		const body =
			shape.shape === 'circle'
				? `<circle cx="5" cy="5" r="4" fill="${color}"/>`
				: `<path d="${escapeSvgAttr(shape.d ?? '')}" fill="${shape.strokeOnly ? 'none' : color}"${shape.strokeOnly ? ` stroke="${color}"` : ''}/>`;
		markup.push(
			`<marker id="${id}" viewBox="0 0 10 10" refX="${shape.refX}" refY="5" markerWidth="${shape.markerWidth}" markerHeight="${shape.markerHeight}" markerUnits="strokeWidth" orient="auto-start-reverse" preserveAspectRatio="none">${body}</marker>`,
		);
		return `url(#${id})`;
	};
	const start = add('Start');
	const end = add('End');
	return { markup: markup.join(''), start, end };
}
