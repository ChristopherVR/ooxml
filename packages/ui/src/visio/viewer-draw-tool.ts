import type { VisioPage, VisioShape } from 'ooxml-core/visio';
import type { ViewerController } from './controller.js';
import { editErrorMessage, isEditCancellation } from './edit-error.js';

/** Visio snaps new geometry to ruler subdivisions; 1/16 inch matches its default fine grid. */
const SNAP = 1 / 16;
const MIN_SIZE = SNAP;
const snap = (value: number) => Math.round(value / SNAP) * SNAP;

/** Next free numeric shape ID on the page tree. Core still rejects any collision. */
export function nextShapeId(page: VisioPage): string {
	let max = 0;
	const pending: VisioShape[] = [...page.shapes];
	while (pending.length) {
		const shape = pending.pop()!;
		if (/^\d{1,9}$/.test(shape.id)) max = Math.max(max, Number(shape.id));
		pending.push(...shape.children);
	}
	return String(max + 1);
}

/** Page inches with a top-left origin (SVG user space) from a client point. */
export function pagePoint(
	svg: SVGSVGElement,
	page: VisioPage,
	event: Pick<MouseEvent, 'clientX' | 'clientY'>,
) {
	const matrix = svg.getScreenCTM?.();
	if (!matrix) return undefined;
	const point = new DOMPoint(event.clientX, event.clientY).matrixTransform(matrix.inverse());
	return {
		x: snap(Math.max(0, Math.min(page.width, point.x))),
		y: snap(Math.max(0, Math.min(page.height, point.y))),
	};
}

/**
 * Create one rectangle through core and select it. `centre` is in top-left page inches (SVG
 * space); core pins are centre points in bottom-left, y-up drawing inches.
 */
export async function insertRectangle(
	controller: ViewerController,
	page: VisioPage,
	centre: { x: number; y: number },
	size: { width: number; height: number },
): Promise<string> {
	const shapeId = nextShapeId(page);
	const pageId = page.id;
	await controller.applyEdits([
		{
			type: 'create-rectangle',
			pageId,
			shapeId,
			x: centre.x,
			y: page.height - centre.y,
			width: size.width,
			height: size.height,
		},
	]);
	const created = controller.state.document?.pages
		.find((candidate) => candidate.id === pageId)
		?.shapes.find((shape) => shape.id === shapeId);
	if (created) controller.selectShape({ id: shapeId, name: created.name, pageId });
	return shapeId;
}

/** Drag-to-draw rectangle tool. The preview is viewer-only; the shape is created by core. */
export class RectangleDrawTool {
	#drag:
		| {
				pointer: number;
				svg: SVGSVGElement;
				page: VisioPage;
				x: number;
				y: number;
				rect: SVGRectElement;
		  }
		| undefined;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; announce(message: string): void },
	) {}
	get drawing(): boolean {
		return !!this.#drag;
	}
	wire(): () => void {
		const Abort = this.viewport.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal };
		const viewport = this.viewport;
		// Drawing must not also select the shape under the pointer.
		viewport.addEventListener(
			'click',
			(event) => {
				if (this.options.active()) event.stopImmediatePropagation();
			},
			{ ...options, capture: true },
		);
		viewport.addEventListener('pointerdown', (event) => this.#start(event), options);
		viewport.addEventListener('pointermove', (event) => this.#move(event), options);
		viewport.addEventListener('pointerup', (event) => void this.#finish(event), options);
		viewport.addEventListener('pointercancel', () => this.#cancel(), options);
		viewport.addEventListener(
			'keydown',
			(event) => {
				if (event.key === 'Escape' && this.#drag) {
					event.preventDefault();
					this.#cancel();
				}
			},
			options,
		);
		return () => {
			++this.#request;
			this.#cancel();
			events.abort();
		};
	}
	#start(event: PointerEvent): void {
		if (!this.options.active() || event.button !== 0 || this.#drag) return;
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		if (!page || !svg || !state.edit.sourceAvailable || state.edit.busy) return;
		const start = pagePoint(svg, page, event);
		if (!start) return;
		event.preventDefault();
		const rect = this.viewport.ownerDocument.createElementNS('http://www.w3.org/2000/svg', 'rect');
		rect.classList.add('draw-preview');
		rect.setAttribute('vector-effect', 'non-scaling-stroke');
		svg.append(rect);
		this.#drag = { pointer: event.pointerId, svg, page, ...start, rect };
		this.viewport.setPointerCapture?.(event.pointerId);
		this.#update(start);
	}
	#move(event: PointerEvent): void {
		if (!this.#drag || event.pointerId !== this.#drag.pointer) return;
		const point = pagePoint(this.#drag.svg, this.#drag.page, event);
		if (point) this.#update(point);
	}
	#update(point: { x: number; y: number }): void {
		const { x, y, rect } = this.#drag!;
		rect.setAttribute('x', String(Math.min(x, point.x)));
		rect.setAttribute('y', String(Math.min(y, point.y)));
		rect.setAttribute('width', String(Math.abs(point.x - x)));
		rect.setAttribute('height', String(Math.abs(point.y - y)));
	}
	#cancel(): void {
		this.#drag?.rect.remove();
		this.#drag = undefined;
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const end = pagePoint(drag.svg, drag.page, event) ?? drag;
		this.#cancel();
		const width = Math.abs(end.x - drag.x),
			height = Math.abs(end.y - drag.y);
		if (width < MIN_SIZE || height < MIN_SIZE) {
			this.options.announce('Drag on the page to draw a rectangle.');
			return;
		}
		const state = this.controller.state;
		if (state.document?.pages[state.pageIndex] !== drag.page) return;
		const request = ++this.#request;
		try {
			const shapeId = await insertRectangle(
				this.controller,
				drag.page,
				{ x: (drag.x + end.x) / 2, y: (drag.y + end.y) / 2 },
				{ width, height },
			);
			if (request !== this.#request) return;
			this.options.announce(
				`Rectangle ${shapeId} added (${+width.toFixed(4)} × ${+height.toFixed(4)} in).`,
			);
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		}
	}
}
