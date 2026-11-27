import {
	visioOpenArrowPath,
	visioFilledArrow,
	layoutVisioFilledArrowLine,
	type VisioStyle,
} from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg';
let markerId = 0;
/** Codes drawn by the generic marker when their native glyph cannot be laid out. */
const GENERIC = new Set([2, 4, 5, 6]);
/** Native measured arrow subsets with explicit filled-path fallbacks. */
export function applyArrowheads(
	path: SVGPathElement,
	style: VisioStyle,
	defs: SVGDefsElement,
	warnings: Set<string>,
): void {
	const filledStart = visioFilledArrow(
		style.startArrow,
		style.startArrowSize ?? 2,
		style.lineWidth,
	);
	const filledEnd = visioFilledArrow(style.endArrow, style.endArrowSize ?? 2, style.lineWidth);
	const layout =
		(filledStart || filledEnd) && (style.lineCap === undefined || style.lineCap === 'round')
			? layoutVisioFilledArrowLine(path.getAttribute('d') ?? '', filledStart, filledEnd)
			: undefined;
	if (layout) path.setAttribute('d', layout.path);
	for (const side of ['start', 'end'] as const) {
		const code = side === 'start' ? style.startArrow : style.endArrow;
		if (!code) continue;
		const size = side === 'start' ? (style.startArrowSize ?? 2) : (style.endArrowSize ?? 2);
		const filledArrow = side === 'start' ? filledStart : filledEnd;
		const openPath = visioOpenArrowPath(code, size, style.lineWidth);
		// Codes 1 to 45 are Visio's built-in line ends; others name a custom (USE) line end.
		if (!filledArrow && !openPath && !GENERIC.has(code)) {
			warnings.add(`Arrowhead style ${code} is not rendered in this build.`);
			continue;
		}
		const id = `visio-arrow-${++markerId}`;
		const marker = svgElement('marker');
		marker.id = id;
		// Off a straight round-capped line the glyph keeps its shape but the line is not set back.
		const setback = layout ? (side === 'start' ? layout.startSetback : layout.endSetback) : 0;
		if (filledArrow && (layout || !GENERIC.has(code))) {
			if (!layout)
				warnings.add('Arrowheads on curved or square-capped lines are not set back from the end.');
			marker.setAttribute('viewBox', '-1 -1 2 2');
			marker.setAttribute('refX', String(-setback));
			marker.setAttribute('refY', '0');
			marker.setAttribute('markerWidth', '2');
			marker.setAttribute('markerHeight', '2');
			marker.setAttribute('markerUnits', 'userSpaceOnUse');
			marker.setAttribute('orient', 'auto-start-reverse');
			marker.setAttribute('overflow', 'visible');
			const glyph = svgElement('path');
			glyph.setAttribute('d', filledArrow.path);
			if (filledArrow.filled) {
				glyph.setAttribute('fill', safeColor(style.lineColor));
				glyph.setAttribute('stroke', 'none');
			} else {
				// A closed outline (hollow triangle, ring, diamond) drawn with the line's own stroke.
				glyph.setAttribute('fill', 'none');
				glyph.setAttribute('stroke', safeColor(style.lineColor));
				glyph.setAttribute('stroke-width', String(style.lineWidth));
				glyph.setAttribute('stroke-linecap', 'round');
				glyph.setAttribute('stroke-linejoin', 'round');
			}
			glyph.setAttribute('opacity', String(style.lineOpacity));
			marker.append(glyph);
			defs.append(marker);
			path.setAttribute(`marker-${side}`, `url(#${id})`);
			continue;
		}
		if (code === 6) {
			warnings.add(
				'Arrowhead style 6 requires a straight round-capped line with non-overlapping filled setbacks.',
			);
			continue;
		}
		if (openPath) {
			// These native open styles have no setback and use round caps
			// independent of the connector's line cap.
			marker.setAttribute('viewBox', '-1 -1 2 2');
			marker.setAttribute('refX', '0');
			marker.setAttribute('refY', '0');
			marker.setAttribute('markerWidth', '2');
			marker.setAttribute('markerHeight', '2');
			marker.setAttribute('markerUnits', 'userSpaceOnUse');
			marker.setAttribute('orient', 'auto-start-reverse');
			marker.setAttribute('overflow', 'visible');
			const tick = svgElement('path');
			tick.setAttribute('d', openPath);
			tick.setAttribute('fill', 'none');
			tick.setAttribute('stroke', safeColor(style.lineColor));
			tick.setAttribute('stroke-width', String(style.lineWidth));
			tick.setAttribute('stroke-linecap', 'round');
			tick.setAttribute('stroke-linejoin', 'round');
			tick.setAttribute('opacity', String(style.lineOpacity));
			marker.append(tick);
			defs.append(marker);
			path.setAttribute(`marker-${side}`, `url(#${id})`);
			continue;
		}
		// Include the code-5 outline inside its viewport instead of clipping the base.
		marker.setAttribute('viewBox', code === 5 ? '-1 -5 11 10' : '0 -4 10 8');
		marker.setAttribute('refX', '9');
		marker.setAttribute('refY', '0');
		marker.setAttribute('orient', 'auto-start-reverse');
		marker.setAttribute('markerUnits', 'userSpaceOnUse');
		const length = (0.1 + Math.max(0, Math.min(6, size)) * 0.045) * (code < 3 ? 0.85 : 1);
		marker.setAttribute('markerWidth', String(length));
		marker.setAttribute('markerHeight', String(length * 0.8));
		const glyph = svgElement('path');
		const filled = code === 2 || code === 4 || code === 5;
		// MS-VSDX BeginArrow code 5 specifies a triangle with an inward-curved base.
		// This original bounded glyph models that topology; its proportions remain approximate.
		glyph.setAttribute(
			'd',
			code === 5
				? 'M 0 -4 L 9 0 L 0 4 Q 4 0 0 -4 Z'
				: filled
					? 'M 0 -4 L 9 0 L 0 4 Z'
					: 'M 0 -4 L 9 0 L 0 4',
		);
		glyph.setAttribute('fill', filled ? safeColor(style.lineColor) : 'none');
		glyph.setAttribute('stroke', safeColor(style.lineColor));
		glyph.setAttribute('stroke-width', '1');
		if (code === 5) glyph.setAttribute('stroke-linejoin', 'round');
		glyph.setAttribute('opacity', String(style.lineOpacity));
		marker.append(glyph);
		defs.append(marker);
		path.setAttribute(`marker-${side}`, `url(#${id})`);
		warnings.add('Common arrowhead shapes are rendered; arrowhead sizing is approximate.');
	}
}
