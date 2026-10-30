import type { VisioImage, VisioShape } from 'ooxml-core/visio';
import type { RenderResources } from './render-resources';
import { svgElement } from './render-svg';

let clipId = 0;
let filterId = 0;

/**
 * Picture Format > Brightness and Contrast as one colour transfer per channel: contrast scales
 * around mid grey, brightness shifts. Both are 0 to 1 with 0.5 neutral, as Visio stores them.
 */
function adjustmentFilter(image: VisioImage): SVGFilterElement | undefined {
	const brightness = image.brightness ?? 0.5,
		contrast = image.contrast ?? 0.5;
	if (brightness === 0.5 && contrast === 0.5) return undefined;
	const slope = 1 + (contrast - 0.5) * 2;
	const intercept = 0.5 * (1 - slope) + (brightness - 0.5);
	const filter = svgElement('filter') as SVGFilterElement;
	filter.id = `visio-image-adjust-${++filterId}`;
	filter.setAttribute('color-interpolation-filters', 'sRGB');
	const transfer = svgElement('feComponentTransfer');
	for (const channel of ['feFuncR', 'feFuncG', 'feFuncB'] as const) {
		const func = svgElement(channel);
		func.setAttribute('type', 'linear');
		func.setAttribute('slope', String(slope));
		func.setAttribute('intercept', String(intercept));
		transfer.append(func);
	}
	filter.append(transfer);
	return filter;
}

/** URLs are created only from bounded embedded raster bytes, never from document strings. */
export function renderImage(
	shape: VisioShape,
	resources: RenderResources,
): SVGGElement | undefined {
	if (!shape.image) return undefined;
	const image = shape.image;
	const url = resources.imageUrl(image);
	const node = resources.portable ? svgElement('use') : svgElement('image');
	const width = image.width ?? shape.width,
		height = image.height ?? shape.height;
	const x = image.x ?? 0,
		y = image.y ?? 0;
	if (resources.portable) node.setAttributeNS('http://www.w3.org/1999/xlink', 'xlink:href', url);
	else node.setAttribute('href', url);
	node.setAttribute('width', String(width));
	node.setAttribute('height', String(height));
	node.setAttribute('transform', `translate(${x} ${y + height}) scale(1 -1)`);
	node.setAttribute('preserveAspectRatio', 'none');
	node.setAttribute('opacity', String(image.opacity ?? 1));
	const group = svgElement('g'),
		clip = svgElement('clipPath'),
		rectangle = svgElement('rect');
	clip.id = `visio-image-clip-${++clipId}`;
	clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
	rectangle.setAttribute('width', String(shape.width));
	rectangle.setAttribute('height', String(shape.height));
	clip.append(rectangle);
	const defs = svgElement('defs');
	defs.append(clip);
	const filter = adjustmentFilter(image);
	if (filter) {
		defs.append(filter);
		node.setAttribute('filter', `url(#${filter.id})`);
	}
	group.append(defs);
	// Clip on an untransformed wrapper, so clipping remains in the shape's frame.
	const cropped = svgElement('g');
	cropped.setAttribute('clip-path', `url(#${clip.id})`);
	cropped.append(node);
	group.append(cropped);
	return group;
}
