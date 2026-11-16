/**
 * Approximate SmartArt 3D in the shared SVG painter (Word and Excel): a shape's bevel is lit with
 * an SVG specular lighting filter from its light rig's direction, its extrusion is drawn as
 * stacked, darkened copies behind the face along the camera's depth direction, and a diagram's
 * spatial camera (a Scene quick style) projects each shape's frame through the fitted scene camera
 * of `ooxml-core/diagram` (an affine map per shape, so perspective is approximated piecewise).
 * Nothing here is a 3D renderer: no hidden-surface removal, no per-pixel material model, no
 * contour, and text is turned with its shape. PowerPoint's lit scene is a separate renderer.
 */
import {
	diagramShape3dView,
	projectDiagramScenePoint,
	resolveDiagramSceneCamera,
	type DiagramDrawingShape,
	type DiagramSceneCamera,
} from 'ooxml-core/diagram';
import type { DrawingScene3D } from 'ooxml-core/drawingml';

const NS = 'http://www.w3.org/2000/svg';
const EMU_PER_PX = 9525;
/** The ECMA-376 default bevel width and height (6pt). */
const DEFAULT_BEVEL_EMU = 76200;
/** Extrusion copies drawn between the face and the back. */
const DEPTH_STEPS = 6;
let serial = 0;

/** SVG `feDistantLight` azimuth (clockwise from +x, y down) of a light rig direction. */
const AZIMUTH: Readonly<Record<string, number>> = {
	t: 270,
	tr: 315,
	r: 0,
	br: 45,
	b: 90,
	bl: 135,
	l: 180,
	tl: 225,
};

/** Specular strength of a preset material: a hint, not Office's material model. */
const SPECULAR: Readonly<Record<string, number>> = {
	metal: 1.4,
	softmetal: 1.1,
	clear: 1.2,
	plastic: 1,
	warmMatte: 0.55,
	matte: 0.45,
	powder: 0.35,
	translucentPowder: 0.35,
	softEdge: 0.7,
	dkEdge: 0.8,
	flat: 0,
};

/** A shape placed through the diagram camera: the affine map of its frame and the scene around it. */
export interface SmartArt3dPlacement {
	camera: DiagramSceneCamera;
	/** The diagram centre the camera turns about, in layout px. */
	centre: { x: number; y: number };
	/** `matrix(a b c d e f)` values mapping the shape's own coordinates onto the screen. */
	matrix: readonly number[];
}

/**
 * The affine placement (`matrix` values) of a shape's frame under a camera: the frame's origin,
 * x and y edges are projected and the shape's own coordinates mapped onto them.
 */
export function projectedPlacement(
	camera: DiagramSceneCamera,
	centre: { x: number; y: number },
	frame: { x: number; y: number; width: number; height: number },
	z = 0,
): number[] {
	const w = frame.width > 0 ? frame.width : 1;
	const h = frame.height > 0 ? frame.height : 1;
	const origin = projectDiagramScenePoint(camera, centre, { x: frame.x, y: frame.y, z });
	const right = projectDiagramScenePoint(camera, centre, { x: frame.x + w, y: frame.y, z });
	const down = projectDiagramScenePoint(camera, centre, { x: frame.x, y: frame.y + h, z });
	const values = [
		(right.x - origin.x) / w,
		(right.y - origin.y) / w,
		(down.x - origin.x) / h,
		(down.y - origin.y) / h,
		origin.x,
		origin.y,
	];
	return values.map((value) => Number(value.toFixed(4)));
}

function bevelFilter(
	doc: Document,
	defs: SVGDefsElement,
	widthPx: number,
	heightPx: number,
	direction: string,
	material: string | undefined,
): string | undefined {
	const specular = SPECULAR[material ?? ''] ?? 0.8;
	if (specular <= 0) return undefined;
	const filter = doc.createElementNS(NS, 'filter');
	filter.id = `office-smartart-bevel-${++serial}`;
	filter.setAttribute('color-interpolation-filters', 'sRGB');
	const blur = doc.createElementNS(NS, 'feGaussianBlur');
	blur.setAttribute('in', 'SourceAlpha');
	blur.setAttribute('stdDeviation', String(Math.max(0.5, widthPx / 2)));
	blur.setAttribute('result', 'bump');
	const light = doc.createElementNS(NS, 'feSpecularLighting');
	light.setAttribute('in', 'bump');
	light.setAttribute('surfaceScale', String(Math.min(10, Math.max(1, (heightPx / widthPx) * 4))));
	light.setAttribute('specularConstant', String(specular));
	light.setAttribute('specularExponent', '18');
	light.setAttribute('lighting-color', '#ffffff');
	light.setAttribute('result', 'highlight');
	const source = doc.createElementNS(NS, 'feDistantLight');
	source.setAttribute('azimuth', String(AZIMUTH[direction] ?? 270));
	source.setAttribute('elevation', '40');
	light.append(source);
	const clip = doc.createElementNS(NS, 'feComposite');
	clip.setAttribute('in', 'highlight');
	clip.setAttribute('in2', 'SourceAlpha');
	clip.setAttribute('operator', 'in');
	clip.setAttribute('result', 'lit');
	const merge = doc.createElementNS(NS, 'feComposite');
	merge.setAttribute('in', 'SourceGraphic');
	merge.setAttribute('in2', 'lit');
	merge.setAttribute('operator', 'arithmetic');
	merge.setAttribute('k1', '0');
	merge.setAttribute('k2', '1');
	merge.setAttribute('k3', '0.6');
	merge.setAttribute('k4', '0');
	filter.append(blur, light, clip, merge);
	defs.append(filter);
	return `url(#${filter.id})`;
}

/**
 * Where the back of an extruded shape lands, in the shape's own coordinates: the camera's depth
 * direction at the shape's centre, brought back through the placement's linear part. Front-on
 * (no camera, or one that looks straight at the face) an extrusion is hidden behind its face.
 */
function depthOffset(
	shape: DiagramDrawingShape,
	scene: DrawingScene3D | undefined,
	placement: SmartArt3dPlacement | undefined,
	depthPx: number,
): { dx: number; dy: number } | undefined {
	const w = shape.frame.width / EMU_PER_PX;
	const h = shape.frame.height / EMU_PER_PX;
	let dx: number;
	let dy: number;
	if (placement) {
		const centre = { x: shape.frame.x / EMU_PER_PX + w / 2, y: shape.frame.y / EMU_PER_PX + h / 2 };
		const face = projectDiagramScenePoint(placement.camera, placement.centre, { ...centre, z: 0 });
		const back = projectDiagramScenePoint(placement.camera, placement.centre, {
			...centre,
			z: -depthPx,
		});
		const [a = 1, b = 0, c = 0, d = 1] = placement.matrix;
		const det = a * d - b * c;
		if (Math.abs(det) < 1e-6) return undefined;
		const sx = back.x - face.x;
		const sy = back.y - face.y;
		dx = (d * sx - c * sy) / det;
		dy = (-b * sx + a * sy) / det;
	} else {
		const camera = resolveDiagramSceneCamera(scene);
		if (!camera) return undefined;
		const centre = { x: w / 2, y: h / 2 };
		const face = projectDiagramScenePoint(camera, centre, { ...centre, z: 0 });
		const back = projectDiagramScenePoint(camera, centre, { ...centre, z: -depthPx });
		dx = back.x - face.x;
		dy = back.y - face.y;
	}
	return Math.hypot(dx, dy) > 0.25 ? { dx, dy } : undefined;
}

/**
 * Paint a shape's 3D around its already-filled `body` inside `group`. Returns false when the shape
 * carries no 3D the painter can approximate (the caller reports it flattened).
 */
export function paintDiagramShape3d(
	doc: Document,
	defs: SVGDefsElement,
	group: SVGGElement,
	body: SVGElement,
	shape: DiagramDrawingShape,
	placement: SmartArt3dPlacement | undefined,
): boolean {
	const view = diagramShape3dView(shape);
	if (!view) return Boolean(placement);
	const solid = view.shape;
	let painted = Boolean(placement);
	const depthPx = (solid?.extrusionHeight ?? 0) / EMU_PER_PX;
	const offset = depthPx > 0 ? depthOffset(shape, view.scene, placement, depthPx) : undefined;
	if (offset) {
		const { dx, dy } = offset;
		const side = doc.createElementNS(NS, 'g');
		side.dataset.extrusion = '';
		for (let step = DEPTH_STEPS; step >= 1; step--) {
			const copy = body.cloneNode(false) as SVGElement;
			copy.removeAttribute('filter');
			copy.setAttribute(
				'transform',
				`translate(${((dx * step) / DEPTH_STEPS).toFixed(2)} ${((dy * step) / DEPTH_STEPS).toFixed(2)})`,
			);
			side.append(copy);
		}
		const shade = body.cloneNode(false) as SVGElement;
		shade.removeAttribute('filter');
		shade.setAttribute('transform', `translate(${(dx / 2).toFixed(2)} ${(dy / 2).toFixed(2)})`);
		shade.setAttribute('fill', '#000000');
		shade.setAttribute('fill-opacity', '0.35');
		shade.setAttribute('stroke', 'none');
		side.append(shade);
		group.insertBefore(side, body);
		painted = true;
	}
	const bevel =
		solid &&
		(solid.bevelTopType !== undefined || solid.bevelTopWidth || solid.bevelTopHeight) &&
		solid.bevelTopType !== 'none';
	if (bevel) {
		const width = (solid.bevelTopWidth ?? DEFAULT_BEVEL_EMU) / EMU_PER_PX;
		const height = (solid.bevelTopHeight ?? DEFAULT_BEVEL_EMU) / EMU_PER_PX;
		const half = Math.min(shape.frame.width, shape.frame.height) / EMU_PER_PX / 2;
		const filter = bevelFilter(
			doc,
			defs,
			Math.max(0.5, Math.min(width, half * 0.9)),
			Math.max(0.5, height),
			view.scene?.lightRigDirection ?? 't',
			solid.presetMaterial,
		);
		if (filter) {
			const wrap = doc.createElementNS(NS, 'g');
			wrap.setAttribute('filter', filter);
			wrap.dataset.bevel = solid.bevelTopType ?? 'circle';
			group.replaceChild(wrap, body);
			wrap.append(body);
		}
		painted = true;
	}
	return painted || Boolean(solid?.presetMaterial);
}
