import {
	BEVEL_BLUR_WIDTHS,
	bevelHeightTable,
	bevelLight,
	bevelMaterial,
	createProjection,
	projectionMatrix,
	visioRotation3dMatrix,
	type VisioMatrix,
	type VisioShape,
} from 'ooxml-core/visio';
import { svgElement } from './render-svg';

let bevelId = 0;
/** Filter lighting measures heights in device pixels; a nominal 96 per inch (100% zoom). */
const PIXELS_PER_INCH = 96;

const scale = (world: VisioMatrix) =>
	Math.sqrt(Math.abs(world[0] * world[3] - world[1] * world[2])) || 1;

/**
 * Bevel: the shape's filled geometry becomes a height field and is lit.
 *
 * The geometry's alpha is blurred (the blur reads as a smooth distance to the edge), a transfer
 * table from `bevelHeightTable` turns that distance into the bevel type's profile height, and
 * `feDiffuseLighting` lights it with the directional light fitted to the shape's
 * `BevelLightingType` (`bevelLight`). Only the departure from the flat face is kept: black where
 * the surface is darker than the face, white where it is brighter, scaled by the material. The
 * powder materials add a white lift. Straight edges follow the profile; corners are rounded by the
 * blur where Visio mitres them, and the strength changes slightly with zoom because SVG lighting
 * works on device pixels. See `bevel-shading.ts` for what was measured and what was shaped by eye.
 */
export function renderBevel(
	shape: VisioShape,
	paths: readonly SVGPathElement[],
	defs: SVGDefsElement,
	world: VisioMatrix,
	/** The filter region when the paths were projected beyond the flat shape's box. */
	box?: { x: number; y: number; width: number; height: number },
): SVGGElement | undefined {
	const bevel = shape.style.bevel;
	const filled = paths.filter((path) => path.getAttribute('fill') !== 'none');
	// A bevel with no size only makes the shape solid for the extrusion and contour.
	if (!bevel || !filled.length || bevel.width <= 0 || bevel.height <= 0) return undefined;
	const material = bevelMaterial(bevel.material ?? 0);
	if (material.contrast <= 0 && material.lift <= 0) return undefined;
	const unit = scale(world);
	const width = bevel.width / unit;
	const light = bevelLight(bevel.lighting ?? 0);
	const table = bevelHeightTable(bevel.type);
	const filter = svgElement('filter');
	filter.id = `visio-bevel-${++bevelId}`;
	filter.setAttribute('filterUnits', 'userSpaceOnUse');
	const region = box ?? {
		x: Math.min(0, shape.width),
		y: Math.min(0, shape.height),
		width: Math.abs(shape.width),
		height: Math.abs(shape.height),
	};
	filter.setAttribute('x', String(region.x - width));
	filter.setAttribute('y', String(region.y - width));
	filter.setAttribute('width', String(region.width + width * 2));
	filter.setAttribute('height', String(region.height + width * 2));
	const add = (name: keyof SVGElementTagNameMap, attributes: Record<string, string | number>) => {
		const node = svgElement(name);
		for (const [key, value] of Object.entries(attributes)) node.setAttribute(key, String(value));
		filter.append(node);
		return node;
	};
	const merge: string[] = [];
	if (material.lift > 0) {
		add('feFlood', { 'flood-color': '#ffffff', 'flood-opacity': material.lift, result: 'lift' });
		add('feComposite', { in: 'lift', in2: 'SourceAlpha', operator: 'in', result: 'lifted' });
		merge.push('lifted');
	}
	if (material.contrast > 0) {
		add('feGaussianBlur', {
			in: 'SourceAlpha',
			stdDeviation: width * BEVEL_BLUR_WIDTHS || 1e-6,
			result: 'blurred',
		});
		const transfer = add('feComponentTransfer', { in: 'blurred', result: 'height' });
		const func = svgElement('feFuncA');
		func.setAttribute('type', 'table');
		func.setAttribute('tableValues', table.values.join(' '));
		transfer.append(func);
		const lit = add('feDiffuseLighting', {
			in: 'height',
			surfaceScale: Math.min(60, table.rise * bevel.height * PIXELS_PER_INCH),
			diffuseConstant: 1,
			'lighting-color': '#ffffff',
			result: 'lit',
		});
		// The lighting runs in the page's y-down frame, like the fitted azimuth.
		const distant = svgElement('feDistantLight');
		lit.append(distant);
		distant.setAttribute('azimuth', String(light.azimuth));
		distant.setAttribute('elevation', String(Math.max(1, Math.min(90, light.elevation))));
		// Brightness against the face is r = (ambient + diffuse * R) / face, R the lit grey.
		const [offset, slope] = [light.ambient / light.face, light.diffuse / light.face];
		const c = material.contrast;
		const row = (color: number, k: number, k0: number) =>
			`0 0 0 0 ${color}  0 0 0 0 ${color}  0 0 0 0 ${color}  ${k} 0 0 0 ${k0}`;
		// Shades (r < 1): black, alpha c * (1 - r). Highlights (r > 1): white, 0.6 c * (r - 1).
		add('feColorMatrix', {
			in: 'lit',
			type: 'matrix',
			values: row(0, -c * slope, c * (1 - offset)),
			result: 'low',
		});
		add('feColorMatrix', {
			in: 'lit',
			type: 'matrix',
			values: row(1, 0.6 * c * slope, 0.6 * c * (offset - 1)),
			result: 'high',
		});
		merge.push('low', 'high');
	}
	add('feMerge', { result: 'both' }).append(
		...merge.map((name) => {
			const node = svgElement('feMergeNode');
			node.setAttribute('in', name);
			return node;
		}),
	);
	add('feComposite', { in: 'both', in2: 'SourceAlpha', operator: 'in' });
	defs.append(filter);
	const group = svgElement('g');
	group.dataset.bevel = '';
	group.dataset.bevelType = String(bevel.type);
	group.setAttribute('pointer-events', 'none');
	group.setAttribute('aria-hidden', 'true');
	group.setAttribute('filter', `url(#${filter.id})`);
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
 * group. A shape with perspective or an extrusion is drawn through `solidRender` instead: its
 * outline is projected point by point, and this wrapper (`projected`) then carries only what stays
 * flat in that mode (text, pictures), see `flatContent`. Visio's own SVG export ignores the
 * rotation, its raster export draws it.
 */
export function rotation3dWrapper(shape: VisioShape, projected = false): SVGGElement | undefined {
	const rotation = shape.style.rotation3d;
	if (!rotation) return undefined;
	const [a, b, c, d] = visioRotation3dMatrix(rotation);
	const cx = shape.width / 2,
		cy = shape.height / 2;
	const wrapper = svgElement('g');
	wrapper.dataset.rotation3d = '';
	if (rotation.perspective > 0) wrapper.dataset.perspective = '';
	if (projected) wrapper.dataset.solid3d = '';
	if (!projected)
		wrapper.setAttribute(
			'transform',
			`matrix(${a} ${b} ${c} ${d} ${cx - a * cx - c * cy} ${cy - b * cx - d * cy})`,
		);
	return wrapper;
}
/**
 * The container for text and pictures of a shape drawn from projected paths: the pinhole
 * projection of the shape's box fitted by one affine matrix (`projectionMatrix`), so the content
 * follows the perspective's overall foreshortening and shear but not its convergence.
 */
export function flatContent(shape: VisioShape, world: VisioMatrix): SVGGElement {
	const rotation = shape.style.rotation3d!;
	const projection = createProjection(
		rotation,
		{ x: shape.width / 2, y: shape.height / 2 },
		scale(world),
	);
	const wrapper = svgElement('g');
	wrapper.dataset.solidFlat = '';
	wrapper.setAttribute(
		'transform',
		`matrix(${projectionMatrix(projection, shape.width, shape.height).join(' ')})`,
	);
	return wrapper;
}
/**
 * The container for the children of a group that is turned in 3-D: the group's projection fitted
 * by one matrix (exact for a parallel rotation, an affine fit for perspective), so every child is
 * turned with the group about the group's centre. A child's own rotation composes inside it.
 */
export function groupChildrenWrapper(
	shape: VisioShape,
	world: VisioMatrix,
): SVGGElement | undefined {
	const rotation = shape.style.rotation3d;
	if (shape.kind !== 'group' || !rotation) return undefined;
	const projection = createProjection(
		rotation,
		{ x: shape.width / 2, y: shape.height / 2 },
		scale(world),
	);
	const wrapper = svgElement('g');
	wrapper.dataset.group3d = '';
	if (rotation.perspective > 0) wrapper.dataset.group3dPerspective = '';
	wrapper.setAttribute(
		'transform',
		`matrix(${projectionMatrix(projection, shape.width, shape.height).join(' ')})`,
	);
	return wrapper;
}
