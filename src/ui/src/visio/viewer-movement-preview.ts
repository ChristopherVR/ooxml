import type { VisioPage } from 'ooxml-core/visio';
import { visioMovePreviewTransform, type VisioPagePoint } from 'ooxml-core/visio/ui';
import { matrix } from './render-svg';
/** Share existing inert SVG resources; moving a preview never mutates the scene or source. */
export function createMovementPreview(svg: SVGSVGElement, page: VisioPage, ids: readonly string[]) {
	const sources = ids.map((id) =>
		Array.from(svg.querySelectorAll<SVGGElement>('[data-shape-id]')).find(
			(group) => group.dataset.shapeId === id && (group.dataset.pageId ?? page.id) === page.id,
		),
	);
	if (sources.some((source) => !source))
		throw new Error('The selected shapes are no longer displayed.');
	const entries = sources.map((source, index) => {
		const original = source!;
		const clone = original.cloneNode(true) as SVGGElement;
		for (const node of [clone, ...Array.from(clone.querySelectorAll('*'))]) {
			for (const name of ['data-shape-id', 'data-page-id', 'tabindex', 'role', 'aria-pressed'])
				node.removeAttribute(name);
		}
		clone.dataset.movementPreview = '';
		clone.setAttribute('aria-hidden', 'true');
		clone.style.pointerEvents = 'none';
		const visibility = original.style.visibility;
		original.after(clone);
		original.style.visibility = 'hidden';
		return {
			original,
			clone,
			visibility,
			shape: page.shapes.find((shape) => shape.id === ids[index])!,
		};
	});
	return {
		update(delta: VisioPagePoint) {
			for (const entry of entries) {
				const transform = visioMovePreviewTransform(entry.shape, delta);
				if (transform) entry.clone.setAttribute('transform', matrix(transform));
			}
		},
		dispose() {
			for (const entry of entries) {
				entry.clone.remove();
				entry.original.style.visibility = entry.visibility;
			}
		},
	};
}
