import type { VisioStyle } from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg.js';
let gradientId = 0;
export function fillPaint(style: VisioStyle, defs: SVGDefsElement): string {
	if (!style.fillGradient) return safeColor(style.fill, '#fff');
	const gradient = svgElement('linearGradient'),
		paint = style.fillGradient;
	gradient.id = `visio-fill-${++gradientId}`;
	gradient.setAttribute('gradientUnits', 'userSpaceOnUse');
	gradient.setAttribute('x1', String(paint.start[0]));
	gradient.setAttribute('y1', String(paint.start[1]));
	gradient.setAttribute('x2', String(paint.end[0]));
	gradient.setAttribute('y2', String(paint.end[1]));
	for (const color of paint.stops) {
		const stop = svgElement('stop');
		stop.setAttribute('offset', String(color.offset));
		stop.setAttribute('stop-color', safeColor(color.color));
		stop.setAttribute('stop-opacity', String(color.opacity));
		gradient.append(stop);
	}
	defs.append(gradient);
	return `url(#${gradient.id})`;
}
