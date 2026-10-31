import {
	createProjection,
	extrudePaths,
	extrusionPathData,
	type Point2,
	type Projection,
	type VisioMatrix,
	type VisioShape,
} from 'ooxml-core/visio';
import { svgElement } from './render-svg';

/** A shape drawn through the point-by-point projection (perspective and/or extrusion). */
export interface SolidRender {
	projection: Projection;
	/** Inches per local unit, so a depth in inches becomes local units. */
	unit: number;
	/** Where the projected shape, its extrusion and its contour can reach, in local units. */
	bounds: { x: number; y: number; width: number; height: number };
}

const inchesPerUnit = (world: VisioMatrix) =>
	Math.sqrt(Math.abs(world[0] * world[3] - world[1] * world[2])) || 1;

/**
 * Perspective cannot be expressed as an SVG matrix, and an extrusion needs the projected outline,
 * so a shape with either is drawn from projected polylines (curves flattened). Without them the
 * 3-D rotation stays a matrix on a wrapper group (`rotation3dWrapper`).
 */
export function solidRender(shape: VisioShape, world: VisioMatrix): SolidRender | undefined {
	const rotation = shape.style.rotation3d;
	if (!rotation || (rotation.perspective <= 0 && !shape.style.extrusion)) return undefined;
	const unit = inchesPerUnit(world);
	const projection = createProjection(rotation, { x: shape.width / 2, y: shape.height / 2 }, unit);
	const depth = (shape.style.extrusion?.depth ?? 0) / unit;
	const margin = ((shape.style.contour?.size ?? 0) + (shape.style.bevel?.width ?? 0)) / unit;
	const corners: Point2[] = [];
	for (const z of depth > 0 ? [0, -depth] : [0])
		for (const x of [0, shape.width])
			for (const y of [0, shape.height]) corners.push(projection.point(x, y, z));
	const xs = corners.map((p) => p.x);
	const ys = corners.map((p) => p.y);
	const [x, y] = [Math.min(...xs) - margin, Math.min(...ys) - margin];
	return {
		projection,
		unit,
		bounds: { x, y, width: Math.max(...xs) + margin - x, height: Math.max(...ys) + margin - y },
	};
}

const clampByte = (value: number) => Math.max(0, Math.min(255, Math.round(value)));
/** Scale a #rrggbb colour: below 1 darkens, above 1 lightens towards white. */
export function shadeColor(hex: string, factor: number): string {
	const channels = [1, 3, 5].map((index) => parseInt(hex.slice(index, index + 2), 16) || 0);
	return `#${channels
		.map((value) => clampByte(factor <= 1 ? value * factor : value + (255 - value) * (factor - 1)))
		.map((value) => value.toString(16).padStart(2, '0'))
		.join('')}`;
}

const layer = (name: 'extrusion' | 'contour'): SVGGElement => {
	const group = svgElement('g');
	group.dataset[name] = '';
	group.setAttribute('pointer-events', 'none');
	group.setAttribute('aria-hidden', 'true');
	return group;
};
const polygon = (d: string, fill: string, stroke?: { color: string; width: number }) => {
	const path = svgElement('path');
	path.setAttribute('d', d);
	path.setAttribute('fill', fill);
	if (stroke) {
		path.setAttribute('stroke', stroke.color);
		path.setAttribute('stroke-width', String(stroke.width));
		path.setAttribute('stroke-linejoin', 'round');
	} else path.setAttribute('stroke', 'none');
	return path;
};

/**
 * The extrusion (back outline, then side faces far to near) and the contour (the extruded outline
 * widened by the contour size, drawn first so only its outer part shows), as one group to put
 * under the shape's own paths. Side shading is fitted to a few of Visio's renders, not its lighting.
 */
export function renderSolid(
	shape: VisioShape,
	solid: SolidRender,
	projectedPaths: readonly string[],
): SVGGElement | undefined {
	const { extrusion, contour } = shape.style;
	if (!extrusion && !contour) return undefined;
	const filled = shape.geometry
		.filter((geometry) => geometry.fill)
		.map((geometry) => geometry.path);
	const faces = extrusion
		? extrudePaths(filled, solid.projection, extrusion.depth / solid.unit)
		: undefined;
	const group = svgElement('g');
	group.dataset.solid = '';
	group.setAttribute('pointer-events', 'none');
	group.setAttribute('aria-hidden', 'true');
	if (contour) {
		const outline = layer('contour');
		const stroke = { color: contour.color, width: (2 * contour.size) / solid.unit };
		for (const d of projectedPaths) outline.append(polygon(d, contour.color, stroke));
		if (faces) {
			if (faces.back.length)
				outline.append(polygon(extrusionPathData(faces.back), contour.color, stroke));
			for (const face of faces.sides)
				outline.append(polygon(extrusionPathData([face.points]), contour.color, stroke));
		}
		group.append(outline);
	}
	if (extrusion && faces) {
		const solidGroup = layer('extrusion');
		if (faces.back.length)
			solidGroup.append(polygon(extrusionPathData(faces.back), shadeColor(extrusion.color, 0.7)));
		for (const face of faces.sides)
			solidGroup.append(
				polygon(
					extrusionPathData([face.points]),
					shadeColor(extrusion.color, face.shade),
					// A hairline in the face colour closes the seams between adjacent faces.
					{ color: shadeColor(extrusion.color, face.shade), width: 0.5 / (solid.unit * 96) },
				),
			);
		group.append(solidGroup);
	}
	return group.childElementCount ? group : undefined;
}
