import type { VisioStyle } from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg.js';
let markerId = 0;
/** Common Visio line/triangle arrow styles. Sizes remain a documented visual approximation. */
export function applyArrowheads(
	path: SVGPathElement,
	style: VisioStyle,
	defs: SVGDefsElement,
	warnings: Set<string>,
): void {
	for (const side of ['start', 'end'] as const) {
		const code = side === 'start' ? style.startArrow : style.endArrow;
		if (!code) continue;
		if (![1, 2, 3, 4].includes(code)) {
			warnings.add(`Arrowhead style ${code} is not rendered in this build.`);
			continue;
		}
		const size = side === 'start' ? (style.startArrowSize ?? 2) : (style.endArrowSize ?? 2);
		const id = `visio-arrow-${++markerId}`;
		const marker = svgElement('marker');
		marker.id = id;
		marker.setAttribute('viewBox', '0 -4 10 8');
		marker.setAttribute('refX', '9');
		marker.setAttribute('refY', '0');
		marker.setAttribute('orient', 'auto-start-reverse');
		marker.setAttribute('markerUnits', 'userSpaceOnUse');
		const length = (0.1 + Math.max(0, Math.min(6, size)) * 0.045) * (code < 3 ? 0.85 : 1);
		marker.setAttribute('markerWidth', String(length));
		marker.setAttribute('markerHeight', String(length * 0.8));
		const glyph = svgElement('path');
		const filled = code === 2 || code === 4;
		glyph.setAttribute('d', filled ? 'M 0 -4 L 9 0 L 0 4 Z' : 'M 0 -4 L 9 0 L 0 4');
		glyph.setAttribute('fill', filled ? safeColor(style.lineColor) : 'none');
		glyph.setAttribute('stroke', safeColor(style.lineColor));
		glyph.setAttribute('stroke-width', '1');
		glyph.setAttribute('opacity', String(style.lineOpacity));
		marker.append(glyph);
		defs.append(marker);
		path.setAttribute(`marker-${side}`, `url(#${id})`);
		warnings.add('Common arrowhead shapes are rendered; arrowhead sizing is approximate.');
	}
}
