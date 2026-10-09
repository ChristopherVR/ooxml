import type { VisioMatrix, VisioShape } from 'ooxml-core/visio';
import { safeColor, svgElement } from './render-svg';

let shadowId = 0;

/**
 * A shape's outer shadow: its geometry painted in the shadow colour, offset in page space and
 * blurred, drawn beneath the shape. Offsets are page inches (Y up); the shape group may be
 * rotated or flipped, so the offset is mapped back through the inverse of its world transform.
 */
export function renderShadow(
	shape: VisioShape,
	paths: readonly SVGPathElement[],
	defs: SVGDefsElement,
	world: VisioMatrix,
): SVGGElement | undefined {
	const shadow = shape.style.shadow;
	if (!shadow || !paths.length || shadow.opacity <= 0) return undefined;
	const [a, b, c, d] = world;
	const det = a * d - b * c;
	if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return undefined;
	const x = (d * shadow.offsetX - c * shadow.offsetY) / det;
	const y = (-b * shadow.offsetX + a * shadow.offsetY) / det;
	const group = svgElement('g');
	group.dataset.shadow = '';
	group.setAttribute('transform', `translate(${x} ${y})`);
	group.setAttribute('opacity', String(shadow.opacity));
	group.setAttribute('pointer-events', 'none');
	group.setAttribute('aria-hidden', 'true');
	const blur = shadow.blur / Math.sqrt(Math.abs(det));
	if (blur > 0) {
		const filter = svgElement('filter');
		filter.id = `visio-shadow-${++shadowId}`;
		// User-space bounds keep blurred shadows of zero-height lines visible.
		const pad = blur * 3 + Math.max(shape.width, shape.height) * 0.05;
		filter.setAttribute('filterUnits', 'userSpaceOnUse');
		filter.setAttribute('x', String(Math.min(0, shape.width) - pad));
		filter.setAttribute('y', String(Math.min(0, shape.height) - pad));
		filter.setAttribute('width', String(Math.abs(shape.width) + pad * 2));
		filter.setAttribute('height', String(Math.abs(shape.height) + pad * 2));
		const gaussian = svgElement('feGaussianBlur');
		gaussian.setAttribute('stdDeviation', String(blur / 2));
		filter.append(gaussian);
		defs.append(filter);
		group.setAttribute('filter', `url(#${filter.id})`);
	}
	const color = safeColor(shadow.color, '#000000');
	for (const source of paths) {
		const path = svgElement('path');
		path.setAttribute('d', source.getAttribute('d') ?? '');
		path.setAttribute('fill', source.getAttribute('fill') === 'none' ? 'none' : color);
		const stroke = source.getAttribute('stroke');
		path.setAttribute('stroke', !stroke || stroke === 'none' ? 'none' : color);
		for (const name of ['stroke-width', 'stroke-linejoin', 'stroke-linecap', 'stroke-dasharray']) {
			const value = source.getAttribute(name);
			if (value !== null) path.setAttribute(name, value);
		}
		group.append(path);
	}
	return group;
}
