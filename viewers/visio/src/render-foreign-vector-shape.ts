import type { VisioShape } from 'ooxml-core/visio';
import type { ForeignVectorBudget } from './foreign-vector-budget.js';
import { renderForeignVector } from './render-foreign-vector.js';
import { svgElement } from './render-svg.js';

let clipId = 0;

/** Core supplies shape-local placement; converter coordinates have a downward Y axis. */
export function renderForeignVectorShape(
	shape: VisioShape,
	budget: ForeignVectorBudget,
): SVGGElement | undefined {
	const foreign = shape.foreignVector;
	if (!foreign) return undefined;
	const node = renderForeignVector(foreign.vector, document, budget);
	node.setAttribute('width', String(foreign.width));
	node.setAttribute('height', String(foreign.height));
	node.setAttribute(
		'transform',
		`translate(${foreign.x} ${foreign.y + foreign.height}) scale(1 -1)`,
	);
	node.setAttribute('opacity', String(foreign.opacity));
	const group = svgElement('g'),
		defs = svgElement('defs'),
		clip = svgElement('clipPath'),
		rectangle = svgElement('rect'),
		cropped = svgElement('g');
	clip.id = `visio-vector-frame-${++clipId}`;
	clip.setAttribute('clipPathUnits', 'userSpaceOnUse');
	rectangle.setAttribute('width', String(shape.width));
	rectangle.setAttribute('height', String(shape.height));
	clip.append(rectangle);
	defs.append(clip);
	// Keep the shape-frame crop outside the placement's Y-axis inversion.
	cropped.setAttribute('clip-path', `url(#${clip.id})`);
	cropped.append(node);
	group.append(defs, cropped);
	return group;
}
