import type { ViewerController, ViewerState } from './controller.js';
interface Controls {
	viewport: HTMLDivElement;
	toolbar: HTMLDivElement;
	pageSelect: HTMLSelectElement;
}
const targetShape = (event: Event) =>
	(event.target as Element)?.closest?.<SVGGElement>('[data-shape-id]');
function selection(target: SVGGElement | undefined | null): ViewerState['selectedShape'] {
	return target
		? {
				id: target.dataset.shapeId!,
				name: target.dataset.shapeName ?? '',
				...(target.dataset.pageId ? { pageId: target.dataset.pageId } : {}),
			}
		: null;
}
/** Own every DOM listener and abort them together when the surface is disposed. */
export function wireViewerInputs(
	controls: Controls,
	controller: ViewerController,
	fit: () => void,
): () => void {
	const { viewport, toolbar, pageSelect } = controls;
	const Abort = viewport.ownerDocument.defaultView?.AbortController ?? AbortController,
		events = new Abort();
	const options = { signal: events.signal };
	pageSelect.addEventListener(
		'change',
		() => controller.setPage(Number(pageSelect.value)),
		options,
	);
	toolbar.addEventListener(
		'click',
		(event) => {
			const action = (event.target as Element)?.closest?.<HTMLButtonElement>('button')?.dataset
				.action;
			if (action === 'in') controller.setZoom(controller.state.zoom * 1.25);
			if (action === 'out') controller.setZoom(controller.state.zoom / 1.25);
			if (action === 'fit') fit();
			if (action === 'actual') controller.setZoom(1);
		},
		options,
	);
	viewport.addEventListener(
		'click',
		(event) => controller.selectShape(selection(targetShape(event))),
		options,
	);
	viewport.addEventListener(
		'keydown',
		(event) => {
			const target = targetShape(event);
			if (
				target &&
				['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft', 'Home', 'End'].includes(event.key)
			) {
				event.preventDefault();
				const shapes = Array.from(viewport.querySelectorAll<SVGGElement>('[data-shape-id]')),
					index = shapes.indexOf(target);
				const next =
					event.key === 'Home'
						? 0
						: event.key === 'End'
							? shapes.length - 1
							: (index +
									(['ArrowRight', 'ArrowDown'].includes(event.key) ? 1 : -1) +
									shapes.length) %
								shapes.length;
				for (const shape of shapes) shape.setAttribute('tabindex', '-1');
				shapes[next]?.setAttribute('tabindex', '0');
				shapes[next]?.focus();
				return;
			}
			if (target && (event.key === 'Enter' || event.key === ' ')) {
				event.preventDefault();
				controller.selectShape(selection(target));
				return;
			}
			if (event.key === 'Escape') controller.selectShape(null);
			if (event.key === '+' || event.key === '=') {
				event.preventDefault();
				controller.setZoom(controller.state.zoom * 1.25);
			}
			if (event.key === '-') {
				event.preventDefault();
				controller.setZoom(controller.state.zoom / 1.25);
			}
			if (event.key === '0') {
				event.preventDefault();
				fit();
			}
		},
		options,
	);
	return () => events.abort();
}
