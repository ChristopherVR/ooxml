import { visioOpenArrowPath, type VisioStyle } from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg.js';
let markerId = 0;
/** Native measured open styles and approximate filled triangle/concave styles. */
export function applyArrowheads(
	path: SVGPathElement,
	style: VisioStyle,
	defs: SVGDefsElement,
	warnings: Set<string>,
): void {
	for (const side of ['start', 'end'] as const) {
		const code = side === 'start' ? style.startArrow : style.endArrow;
		if (!code) continue;
		if (![1, 2, 3, 4, 5, 7, 9].includes(code)) {
			warnings.add(`Arrowhead style ${code} is not rendered in this build.`);
			continue;
		}
		const size = side === 'start' ? (style.startArrowSize ?? 2) : (style.endArrowSize ?? 2);
		const id = `visio-arrow-${++markerId}`;
		const marker = svgElement('marker');
		marker.id = id;
		const openPath = visioOpenArrowPath(code, size, style.lineWidth);
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
