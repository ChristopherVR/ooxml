import type { VisioPage } from 'ooxml-core/visio';
import { VISIO_RESIZE_HANDLES, type VisioResizeFrame } from 'ooxml-core/visio/ui';
import { matrix } from './render-svg';
/** A bounds preview, leaving the saved shape content visible until source acceptance. */
export function createResizeFrame(
	svg: SVGSVGElement,
	page: VisioPage,
	frame: VisioResizeFrame,
	scale: number,
	shapeId: string,
) {
	const create = <K extends keyof SVGElementTagNameMap>(name: K) =>
		svg.ownerDocument.createElementNS('http://www.w3.org/2000/svg', name);
	const element = create('g');
	element.dataset.resizeOverlay = '';
	const rectangle = create('rect');
	rectangle.dataset.resizeFrame = '';
	rectangle.setAttribute('vector-effect', 'non-scaling-stroke');
	element.append(rectangle);
	const handles = VISIO_RESIZE_HANDLES.map((location) => {
		const handle = create('circle');
		handle.dataset.resizeHandle = location.id;
		handle.dataset.resizeShape = shapeId;
		handle.setAttribute('r', String(5 / scale));
		handle.setAttribute('vector-effect', 'non-scaling-stroke');
		handle.setAttribute('aria-label', `Drag ${location.id} resize handle`);
		element.append(handle);
		return { handle, location };
	});
	const update = (frame: VisioResizeFrame) => {
		element.setAttribute(
			'transform',
			`translate(0 ${page.height}) scale(1 -1) ${matrix(frame.transform)}`,
		);
		rectangle.setAttribute('width', String(frame.width));
		rectangle.setAttribute('height', String(frame.height));
		for (const { handle, location } of handles) {
			handle.setAttribute('cx', String(location.x * frame.width));
			handle.setAttribute('cy', String(location.y * frame.height));
		}
	};
	update(frame);
	svg.append(element);
	return { element, rectangle, update, dispose: () => element.remove() };
}
