import { evaluatePresetShape, ooxmlArcToSvg } from 'ooxml-core/geometry';
import type { DiagramDrawingShape } from 'ooxml-core/diagram';
import type { DrawingPath } from 'ooxml-core/drawingml';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** DrawingML arcs share the core preset evaluator's conversion, including full-circle splits. */
export function drawingPathData(path: DrawingPath, width: number, height: number): string {
	const sx = path.width ? width / path.width : 1;
	const sy = path.height ? height / path.height : 1;
	const out: string[] = [];
	let px = 0;
	let py = 0;
	let startX = 0;
	let startY = 0;
	for (const command of path.commands) {
		if (command.op === 'A') {
			const arc = ooxmlArcToSvg(
				command.wR * sx,
				command.hR * sy,
				command.startAngle * 60000,
				command.swingAngle * 60000,
				px,
				py,
			);
			if (arc) {
				out.push(arc.svg);
				px = arc.endX;
				py = arc.endY;
			}
		} else if (command.op === 'Z') {
			out.push('Z');
			px = startX;
			py = startY;
		} else {
			px = command.x * sx;
			py = command.y * sy;
			if (command.op === 'M' || command.op === 'L') {
				out.push(`${command.op}${px} ${py}`);
				if (command.op === 'M') {
					startX = px;
					startY = py;
				}
			} else if (command.op === 'C')
				out.push(
					`C${command.x1 * sx} ${command.y1 * sy} ${command.x2 * sx} ${command.y2 * sy} ${px} ${py}`,
				);
			else if (command.op === 'Q') out.push(`Q${command.x1 * sx} ${command.y1 * sy} ${px} ${py}`);
		}
	}
	return out.join(' ');
}

export function diagramShapeElement(
	doc: Document,
	shape: DiagramDrawingShape,
	approximated: string[],
): SVGElement {
	const width = shape.frame.width / 9525;
	const height = shape.frame.height / 9525;
	if (!shape.paths?.length && ['rect', 'ellipse'].includes(shape.geometry)) {
		const element = doc.createElementNS(SVG_NS, shape.geometry === 'ellipse' ? 'ellipse' : 'rect');
		if (shape.geometry === 'ellipse') {
			element.setAttribute('cx', String(width / 2));
			element.setAttribute('cy', String(height / 2));
			element.setAttribute('rx', String(width / 2));
			element.setAttribute('ry', String(height / 2));
		} else {
			element.setAttribute('width', String(width));
			element.setAttribute('height', String(height));
		}
		return element;
	}
	const evaluated = shape.paths?.length
		? undefined
		: evaluatePresetShape(shape.geometry, width, height, shape.adjustments);
	const paths =
		shape.paths?.map((path) => ({
			d: drawingPathData(path, width, height),
			fill: path.fill,
			stroke: path.stroke,
		})) ?? evaluated?.paths;
	if (paths?.length) {
		const group = doc.createElementNS(SVG_NS, 'g');
		for (const path of paths) {
			const element = doc.createElementNS(SVG_NS, 'path');
			element.setAttribute('d', path.d);
			if (path.fill === 'none') element.setAttribute('fill', 'none');
			if (path.stroke === false) element.setAttribute('stroke', 'none');
			group.append(element);
		}
		return group;
	}
	if (!approximated.includes(shape.geometry)) approximated.push(shape.geometry);
	const rect = doc.createElementNS(SVG_NS, 'rect');
	rect.setAttribute('width', String(width));
	rect.setAttribute('height', String(height));
	rect.setAttribute('data-approximated', shape.geometry);
	return rect;
}
