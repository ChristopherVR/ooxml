import { evaluatePresetShape } from 'ooxml-core/geometry';
import type { DiagramDrawingShape } from 'ooxml-core/diagram';

const SVG_NS = 'http://www.w3.org/2000/svg';

export { drawingPathData } from 'ooxml-core/drawingml';
import { drawingPathData } from 'ooxml-core/drawingml';

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
