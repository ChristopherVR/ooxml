import type { ViewerState } from './controller';
import { pointerShapeTarget } from './viewer-shape-target';
import { hitProbes, hitTolerance } from './viewer-touch-math';

/** A thin connector or line within the finger's reach, when the touch landed beside it. */
export function nearbyShape(
	viewport: HTMLElement,
	svg: SVGSVGElement,
	event: PointerEvent,
	selected: ViewerState['selectedShapes'],
): SVGGElement | null {
	const root = viewport.getRootNode() as Document | ShadowRoot;
	if (typeof root.elementsFromPoint !== 'function') return null;
	for (const probe of hitProbes(hitTolerance(event.pointerType))) {
		const hits = root.elementsFromPoint(event.clientX + probe.x, event.clientY + probe.y);
		for (const element of hits) {
			if (!svg.contains(element) || element.closest('[data-resize-overlay], .rotation-overlay'))
				continue;
			const shape = pointerShapeTarget(element, selected, true);
			if (shape) return shape;
		}
	}
	return null;
}
