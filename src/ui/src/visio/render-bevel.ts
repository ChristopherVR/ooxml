import {
	VISIO_BEVEL_TYPES_INSET,
	VISIO_BEVEL_TYPES_SHARP,
	visioRotation3dMatrix,
	type VisioMatrix,
	type VisioShape,
} from 'ooxml-core/visio';
import { svgElement } from './render-svg';

let bevelId = 0;
/** Flat surfaces light to sin(elevation); the overlay only keeps the difference from that. */
const ELEVATION = 45;
const FLAT = Math.sin((ELEVATION * Math.PI) / 180);

const scale = (world: VisioMatrix) =>
	Math.sqrt(Math.abs(world[0] * world[3] - world[1] * world[2])) || 1;

/**
 * Bevel: an approximation. The shape's filled geometry is blurred by the bevel width, lit from the
 * top left with a diffuse light, and only the departure from a flat surface is kept: white where
 * the surface faces the light, black where it faces away. Inset bevels invert the light and hard
 * profiles narrow the blur. Visio's own bevel profiles, materials and lighting are not reproduced,
 * and the result changes slightly with the zoom because SVG lighting works on device pixels.
 */
export function renderBevel(
	shape: VisioShape,
	paths: readonly SVGPathElement[],
	defs: SVGDefsElement,
	world: VisioMatrix,
): SVGGElement | undefined {
	const bevel = shape.style.bevel;
	const filled = paths.filter((path) => path.getAttribute('fill') !== 'none');
	if (!bevel || !filled.length) return undefined;
	const unit = scale(world);
	const width = Math.max(bevel.width, bevel.height) / unit;
	const sharp = VISIO_BEVEL_TYPES_SHARP.includes(bevel.type);
	const inset = VISIO_BEVEL_TYPES_INSET.includes(bevel.type);
	const radius = (sharp ? width / 4 : width / 2) || 1e-6;
	const filter = svgElement('filter');
	filter.id = `visio-bevel-${++bevelId}`;
	filter.setAttribute('filterUnits', 'userSpaceOnUse');
	filter.setAttribute('x', String(Math.min(0, shape.width) - width));
	filter.setAttribute('y', String(Math.min(0, shape.height) - width));
	filter.setAttribute('width', String(Math.abs(shape.width) + width * 2));
	filter.setAttribute('height', String(Math.abs(shape.height) + width * 2));
	const add = (name: keyof SVGElementTagNameMap, attributes: Record<string, string | number>) => {
		const node = svgElement(name);
		for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
		filter.append(node);
		return node;
	};
	add('feGaussianBlur', { in: 'SourceAlpha', stdDeviation: radius, result: 'height' });
	const light = add('feDiffuseLighting', {
		in: 'height',
		surfaceScale: Math.min(10, 2 + (bevel.height * 72) / 2) * (sharp ? 1.5 : 1),
		diffuseConstant: 1,
		'lighting-color': '#ffffff',
		result: 'lit',
	});
	// The local frame is y-up and the lighting runs in the page's y-down frame.
	const distant = svgElement('feDistantLight');
	light.append(distant);
	distant.setAttribute('azimuth', String(inset ? 45 : 225));
	distant.setAttribute('elevation', String(ELEVATION));
	const gain = sharp ? 5 : 3;
	// Highlights: white with alpha rising above the flat level. Shades: black, below it.
	add('feColorMatrix', {
		in: 'lit',
		type: 'matrix',
		values: `0 0 0 0 1  0 0 0 0 1  0 0 0 0 1  ${gain} 0 0 0 ${-gain * FLAT}`,
		result: 'high',
	});
	add('feColorMatrix', {
		in: 'lit',
		type: 'matrix',
		values: `0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  ${-gain} 0 0 0 ${gain * FLAT}`,
		result: 'low',
	});
	add('feMerge', { result: 'both' }).append(
		...['low', 'high'].map((name) => {
			const node = svgElement('feMergeNode');
			node.setAttribute('in', name);
			return node;
		}),
	);
	add('feComposite', { in: 'both', in2: 'SourceAlpha', operator: 'in' });
	defs.append(filter);
	const group = svgElement('g');
	group.dataset.bevel = '';
	group.setAttribute('pointer-events', 'none');
	group.setAttribute('aria-hidden', 'true');
	group.setAttribute('filter', `url(#${filter.id})`);
	group.setAttribute('opacity', '0.6');
	for (const source of filled) {
		const copy = svgElement('path');
		copy.setAttribute('d', source.getAttribute('d') ?? '');
		copy.setAttribute('fill', '#000000');
		copy.setAttribute('stroke', 'none');
		group.append(copy);
	}
	return group;
}

/**
 * 3-D rotation: a parallel projection of the flat shape about its centre, as a matrix on a wrapper
 * group. Perspective foreshortening and extrusion are not drawn; Visio's own SVG export ignores
 * the rotation, its raster export draws it.
 */
export function rotation3dWrapper(shape: VisioShape): SVGGElement | undefined {
	const rotation = shape.style.rotation3d;
	if (!rotation) return undefined;
	const [a, b, c, d] = visioRotation3dMatrix(rotation);
	const cx = shape.width / 2,
		cy = shape.height / 2;
	const wrapper = svgElement('g');
	wrapper.dataset.rotation3d = '';
	if (rotation.perspective > 0) wrapper.dataset.perspective = '';
	wrapper.setAttribute(
		'transform',
		`matrix(${a} ${b} ${c} ${d} ${cx - a * cx - c * cy} ${cy - b * cx - d * cy})`,
	);
	return wrapper;
}
