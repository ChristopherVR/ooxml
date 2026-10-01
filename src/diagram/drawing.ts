// Reads the cached drawing part of a diagram (`dsp:drawing`): the shape tree the producing
// application last laid out, as a renderer-neutral model in EMU. Generalises
// `parseSmartArtDrawingShapes`/`parseDrawingShape` of the pptx runtime without its theme and
// presentation dependencies (colours stay unresolved; see `resolveDrawingColor`).
import { parseXml, type XmlDocument } from '../xml/index.js';
import { parseStyleReferences } from './definitions.js';
import { parseDrawingFill, parseDrawingLine } from './drawing-fill.js';
import { parseDrawingGeometry, parseFrame, parseTransform } from './drawing-geometry.js';
import { parseDrawingTextBody } from './drawing-text.js';
import { NS, first, stringAttribute, type XmlElement } from './dom.js';
import type { DiagramDrawing, DiagramDrawingShape, DiagramIssue } from './types.js';

interface GroupTransform {
	offsetX: number;
	offsetY: number;
	scaleX: number;
	scaleY: number;
}
const IDENTITY: GroupTransform = { offsetX: 0, offsetY: 0, scaleX: 1, scaleY: 1 };

/** Maps a child's frame through a group's `chOff/chExt` -> `off/ext` mapping. */
function applyGroup(
	frame: { x: number; y: number; width: number; height: number },
	group: GroupTransform,
) {
	return {
		x: Math.round(group.offsetX + frame.x * group.scaleX),
		y: Math.round(group.offsetY + frame.y * group.scaleY),
		width: Math.round(frame.width * group.scaleX),
		height: Math.round(frame.height * group.scaleY),
	};
}

function groupTransform(group: XmlElement, parent: GroupTransform): GroupTransform {
	const xfrm = first(first(group, 'grpSpPr', NS.dsp), 'xfrm', NS.a);
	const off = first(xfrm, 'off', NS.a);
	const ext = first(xfrm, 'ext', NS.a);
	const chOff = first(xfrm, 'chOff', NS.a);
	const chExt = first(xfrm, 'chExt', NS.a);
	const num = (element: XmlElement | undefined, name: string) =>
		Number.parseInt(element?.getAttribute(name) ?? '', 10);
	const cx = num(ext, 'cx');
	const cy = num(ext, 'cy');
	const chCx = num(chExt, 'cx');
	const chCy = num(chExt, 'cy');
	if (!off || ![cx, cy, chCx, chCy].every(Number.isFinite) || chCx === 0 || chCy === 0)
		return parent;
	const scaleX = cx / chCx;
	const scaleY = cy / chCy;
	const local: GroupTransform = {
		offsetX: num(off, 'x') - (num(chOff, 'x') || 0) * scaleX,
		offsetY: num(off, 'y') - (num(chOff, 'y') || 0) * scaleY,
		scaleX,
		scaleY,
	};
	// Compose with the enclosing group.
	return {
		offsetX: parent.offsetX + local.offsetX * parent.scaleX,
		offsetY: parent.offsetY + local.offsetY * parent.scaleY,
		scaleX: parent.scaleX * local.scaleX,
		scaleY: parent.scaleY * local.scaleY,
	};
}

function parseShape(
	sp: XmlElement,
	index: number,
	group: GroupTransform,
	issues: DiagramIssue[],
): DiagramDrawingShape | undefined {
	const spPr = first(sp, 'spPr', NS.dsp);
	const transform = parseTransform(first(spPr, 'xfrm', NS.a));
	if (!transform) return undefined;
	const modelId =
		stringAttribute(sp, 'modelId') ??
		stringAttribute(first(first(sp, 'nvSpPr', NS.dsp), 'cNvPr', NS.dsp), 'id') ??
		`dsp-${index}`;
	const { geometry, adjustments, paths, issue } = parseDrawingGeometry(spPr);
	if (issue) issues.push({ code: issue, message: `Shape ${modelId}: ${issue}.` });
	// A degenerate frame is a stale producer extent, except for `line` geometry, whose rails and
	// stems (Timeline and others) are cached with a zero height or width by design.
	const frame = applyGroup(transform.frame, group);
	if ((frame.width <= 0 || frame.height <= 0) && geometry !== 'line') return undefined;

	const shape: DiagramDrawingShape = {
		modelId,
		frame,
		geometry,
		has3d: Boolean(first(spPr, 'scene3d', NS.a) || first(spPr, 'sp3d', NS.a)),
	};
	if (transform.rotation !== undefined) shape.rotation = transform.rotation;
	if (transform.flipHorizontal) shape.flipHorizontal = true;
	if (transform.flipVertical) shape.flipVertical = true;
	if (adjustments) shape.adjustments = adjustments;
	if (paths) shape.paths = paths;
	const fill = parseDrawingFill(spPr);
	if (fill) shape.fill = fill;
	const line = parseDrawingLine(first(spPr, 'ln', NS.a));
	if (line) shape.line = line;
	const style = first(sp, 'style', NS.dsp);
	if (style) shape.style = parseStyleReferences(style);
	const text = parseDrawingTextBody(first(sp, 'txBody', NS.dsp));
	if (text) shape.text = text;
	const textFrame = parseFrame(first(sp, 'txXfrm', NS.dsp));
	if (textFrame) shape.textFrame = applyGroup(textFrame, group);
	return shape;
}

/** Parses a drawing part. Malformed XML throws; unsupported content is reported in `issues`. */
export function parseDiagramDrawing(source: string | XmlDocument, part?: string): DiagramDrawing {
	const root = (typeof source === 'string' ? parseXml(source, { label: 'DiagramML' }) : source)
		.documentElement;
	const issues: DiagramIssue[] = [];
	const shapes: DiagramDrawingShape[] = [];
	const tree = root.localName === 'spTree' ? root : first(root, 'spTree', NS.dsp);
	let counter = 0;
	const visit = (container: XmlElement, group: GroupTransform) => {
		for (const child of Array.from(container.childNodes)) {
			if (child.nodeType !== 1) continue;
			const element = child as XmlElement;
			if (element.localName === 'sp') {
				const shape = parseShape(element, counter++, group, issues);
				if (shape) shapes.push(shape);
			} else if (element.localName === 'grpSp') visit(element, groupTransform(element, group));
			else if (element.localName === 'pic')
				issues.push({
					code: 'DIAGRAM_DRAWING_PICTURE_SKIPPED',
					message: 'A dsp:pic in the cached drawing is not modelled; the part still holds it.',
				});
		}
	};
	if (tree) visit(tree, IDENTITY);
	else
		issues.push({ code: 'DIAGRAM_DRAWING_EMPTY', message: 'The drawing part has no dsp:spTree.' });
	const withPart = part ? issues.map((issue) => ({ ...issue, part })) : issues;
	return { shapes, issues: withPart };
}

/** Whether the shapes of a drawing include any with 3D scene or shape definitions. */
export const drawingHas3d = (drawing: DiagramDrawing): boolean =>
	drawing.shapes.some((shape) => shape.has3d);
