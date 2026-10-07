import type { VisioStyle } from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg';
import type { RenderResources } from './render-resources';
let gradientId = 0;
export function fillPaint(
	style: VisioStyle,
	defs: SVGDefsElement,
	resources: RenderResources,
): string {
	if (style.fillPattern) {
		const paint = style.fillPattern,
			pattern = svgElement('pattern');
		pattern.id = `visio-fill-${++gradientId}`;
		pattern.setAttribute('patternUnits', 'userSpaceOnUse');
		pattern.setAttribute('width', String(paint.width));
		pattern.setAttribute('height', String(paint.height));
		pattern.setAttribute('viewBox', '0 0 64 64');
		pattern.setAttribute('patternTransform', 'scale(1 -1)');
		const node = svgElement(resources.portable ? 'use' : 'image');
		const url = resources.imageUrl(paint);
		if (resources.portable) node.setAttributeNS('http://www.w3.org/1999/xlink', 'xlink:href', url);
		else node.setAttribute('href', url);
		node.setAttribute('width', '64');
		node.setAttribute('height', '64');
		node.setAttribute('image-rendering', 'optimizeSpeed');
		node.setAttribute('preserveAspectRatio', 'none');
		// Flip the native top-down tile back into local y-up shape coordinates.
		pattern.append(node);
		defs.append(pattern);
		return `url(#${pattern.id})`;
	}
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
