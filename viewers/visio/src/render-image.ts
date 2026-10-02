import type { VisioShape } from 'ooxml-core/visio';
import type { RenderResources } from './render-resources.js';
import { svgElement } from './render-svg.js';

let clipId = 0;

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
	group.append(defs);
	// Clip on an untransformed wrapper, so clipping remains in the shape's frame.
	const cropped = svgElement('g');
	cropped.setAttribute('clip-path', `url(#${clip.id})`);
	cropped.append(node);
	group.append(cropped);
	return group;
}
