import type {
	VisioStyle,
	VisioMatrix,
	VisioLinearGradient,
	VisioRadialGradient,
} from 'ooxml-core/visio';
import {
	visioFillPatternTransform,
	visioRenderedGradientStops,
	visioStrokeGradient,
} from 'ooxml-core/visio/ui';
import { safeColor, svgElement, matrix } from './render-svg';
import type { RenderResources } from './render-resources';
let gradientId = 0;
export function fillPaint(
	style: Pick<VisioStyle, 'fill' | 'fillPattern' | 'fillGradient'>,
	defs: SVGDefsElement,
	resources: RenderResources,
	world: VisioMatrix,
	pageHeight: number,
): string {
	if (style.fillPattern) {
		const paint = style.fillPattern,
			pattern = svgElement('pattern');
		pattern.id = `visio-fill-${++gradientId}`;
		pattern.setAttribute('patternUnits', 'userSpaceOnUse');
		pattern.setAttribute('width', String(paint.width));
		pattern.setAttribute('height', String(paint.height));
		pattern.setAttribute('viewBox', '0 0 64 64');
		const transform = visioFillPatternTransform(world, pageHeight);
		if (transform) pattern.setAttribute('patternTransform', matrix(transform));
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
	const paint = style.fillGradient;
	if (paint.type === 'regions') {
		const stops = visioRenderedGradientStops(paint);
		const pattern = svgElement('pattern');
		pattern.id = `visio-fill-${++gradientId}`;
		pattern.setAttribute('patternUnits', 'objectBoundingBox');
		pattern.setAttribute('patternContentUnits', 'objectBoundingBox');
		pattern.setAttribute('width', '1');
		pattern.setAttribute('height', '1');
		const shapeCoordinates = paint.coordinateSpace === 'shape';
		if (!shapeCoordinates) pattern.setAttribute('patternTransform', 'scale(1 -1)');
		for (const region of paint.regions) {
			const path = svgElement('path');
			// Shared triangle edges must not expose antialiased transparent seams.
			path.setAttribute('shape-rendering', 'crispEdges');
			path.setAttribute(
				'd',
				region.points
					.map(
						(point, index) =>
							`${index ? 'L' : 'M'} ${point[0]} ${shapeCoordinates ? point[1] : 1 - point[1]}`,
					)
					.join(' ') + ' z',
			);
			path.setAttribute(
				'fill',
				gradientPaint(
					{
						type: 'linear',
						start: shapeCoordinates ? region.start : [0, 0],
						end: shapeCoordinates ? region.end : [1, 0],
						stops,
					},
					defs,
					!shapeCoordinates,
					shapeCoordinates ? undefined : region.angle,
				),
			);
			pattern.append(path);
		}
		defs.append(pattern);
		return `url(#${pattern.id})`;
	}
	return gradientPaint(paint, defs);
}

export function linePaint(
	style: VisioStyle,
	defs: SVGDefsElement,
	resources: RenderResources,
	world: VisioMatrix,
	pageHeight: number,
	size: readonly [number, number],
): string {
	const gradient = visioStrokeGradient(style.lineGradient, style.lineWidth, size);
	if (!gradient) return safeColor(style.lineColor);
	return fillPaint(
		{ fill: style.lineColor, fillGradient: gradient },
		defs,
		resources,
		world,
		pageHeight,
	);
}

/** All gradient kinds share the existing stop, color and alpha serialization. */
function gradientPaint(
	paint: VisioLinearGradient | VisioRadialGradient,
	defs: SVGDefsElement,
	normalized = false,
	rotation?: number,
): string {
	const gradient = svgElement(paint.type === 'radial' ? 'radialGradient' : 'linearGradient');
	gradient.id = `visio-fill-${++gradientId}`;
	if (paint.type === 'radial') {
		gradient.setAttribute(
			'gradientUnits',
			paint.coordinateSpace === 'local' ? 'userSpaceOnUse' : 'objectBoundingBox',
		);
		gradient.setAttribute('cx', String(paint.center[0]));
		gradient.setAttribute('cy', String(paint.center[1]));
		gradient.setAttribute('r', String(paint.radius));
	} else {
		normalized ||= paint.boundingBoxAngle !== undefined;
		rotation ??= paint.boundingBoxAngle;
		gradient.setAttribute('gradientUnits', normalized ? 'objectBoundingBox' : 'userSpaceOnUse');
		gradient.setAttribute('x1', String(paint.start[0]));
		gradient.setAttribute('y1', String(paint.start[1]));
		gradient.setAttribute('x2', String(paint.end[0]));
		gradient.setAttribute('y2', String(paint.end[1]));
		if (rotation !== undefined)
			gradient.setAttribute('gradientTransform', `rotate(${rotation} 0.5 0.5)`);
	}
	for (const color of visioRenderedGradientStops(paint)) {
		const stop = svgElement('stop');
		stop.setAttribute('offset', String(color.offset));
		stop.setAttribute('stop-color', safeColor(color.color));
		stop.setAttribute('stop-opacity', String(color.opacity));
		gradient.append(stop);
	}
	defs.append(gradient);
	return `url(#${gradient.id})`;
}
