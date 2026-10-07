import type { VisioPage, VisioGeometryEdit } from 'ooxml-core/visio';
import type { ViewerController } from './controller';
import {
	editErrorMessage,
	isEditCancellation,
	visioPageEditToDrawing,
	visioNextShapeId as nextShapeId,
} from 'ooxml-core/visio/ui';
export { visioNextShapeId as nextShapeId } from 'ooxml-core/visio/ui';

/** Visio snaps new geometry to ruler subdivisions; 1/16 inch matches its default fine grid. */
const SNAP = 1 / 16;
const MIN_SIZE = SNAP;
const snap = (value: number) => Math.round(value / SNAP) * SNAP;

/** Page inches with a top-left origin (SVG user space) from a client point. */
export function pagePoint(
	svg: SVGSVGElement,
	page: VisioPage,
	event: Pick<MouseEvent, 'clientX' | 'clientY'>,
	options: { snap?: boolean; bounded?: boolean } = {},
) {
	const matrix = svg.getScreenCTM?.();
	if (!matrix) return undefined;
	const point = new DOMPoint(event.clientX, event.clientY).matrixTransform(matrix.inverse());
	const coordinate = (value: number, maximum: number) => {
		const bounded = options.bounded === false ? value : Math.max(0, Math.min(maximum, value));
		return options.snap === false ? bounded : snap(bounded);
	};
	return { x: coordinate(point.x, page.width), y: coordinate(point.y, page.height) };
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
	return insertGeometry(controller, page, {
		type: 'create-rectangle',
		pageId: page.id,
		shapeId,
		x: centre.x,
		y: page.height - centre.y,
		width: size.width,
		height: size.height,
	});
}

export async function insertLine(
	controller: ViewerController,
	page: VisioPage,
	begin: { x: number; y: number },
	end: { x: number; y: number },
): Promise<string> {
	return insertGeometry(controller, page, {
		type: 'create-line',
		pageId: page.id,
		shapeId: nextShapeId(page),
		beginX: begin.x,
		beginY: page.height - begin.y,
		endX: end.x,
		endY: page.height - end.y,
	});
}

async function insertGeometry(
	controller: ViewerController,
	page: VisioPage,
	command: VisioGeometryEdit,
): Promise<string> {
	const { shapeId, pageId } = command;
	await controller.applyEdits([visioPageEditToDrawing(page, command)]);
	const created = controller.state.document?.pages
		.find((candidate) => candidate.id === pageId)
		?.shapes.find((shape) => shape.id === shapeId);
	if (created) controller.selectShape({ id: shapeId, name: created.name, pageId });
	return shapeId;
}

/** Shared rectangle/line gesture lifecycle. Preview is DOM-only; core owns creation. */
export class ShapeDrawTool {
	#drag:
		| {
				pointer: number;
				svg: SVGSVGElement;
				page: VisioPage;
				x: number;
				y: number;
				rect: SVGRectElement | SVGLineElement;
				kind: 'rectangle' | 'line';
		  }
		| undefined;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: {
			tool(): 'rectangle' | 'line' | undefined;
			announce(message: string): void;
		},
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
				if (this.options.tool()) event.stopImmediatePropagation();
			},
			{ ...options, capture: true },
		);
		viewport.addEventListener('pointerdown', (event) => this.#start(event), options);
		viewport.addEventListener('pointermove', (event) => this.#move(event), options);
		viewport.addEventListener('pointerup', (event) => void this.#finish(event), options);
		viewport.addEventListener('pointercancel', () => this.#cancel(), options);
		viewport.addEventListener(
			'lostpointercapture',
			(event) => {
				if (event.pointerId === this.#drag?.pointer) this.#cancel();
			},
			options,
		);
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
		const kind = this.options.tool();
		if (!kind || event.button !== 0 || this.#drag) return;
		const state = this.controller.state;
		const page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		if (!page || !svg || !state.edit.sourceAvailable || state.edit.busy) return;
		const start = pagePoint(svg, page, event);
		if (!start) return;
		event.preventDefault();
		const rect = this.viewport.ownerDocument.createElementNS(
			'http://www.w3.org/2000/svg',
			kind === 'line' ? 'line' : 'rect',
		);
		rect.classList.add('draw-preview');
		rect.setAttribute('vector-effect', 'non-scaling-stroke');
		svg.append(rect);
		this.#drag = { pointer: event.pointerId, svg, page, ...start, rect, kind };
		this.viewport.setPointerCapture?.(event.pointerId);
		this.#update(start);
	}
	#move(event: PointerEvent): void {
		if (!this.#drag || event.pointerId !== this.#drag.pointer) return;
		const point = pagePoint(this.#drag.svg, this.#drag.page, event);
		if (point) this.#update(point);
	}
	#update(point: { x: number; y: number }): void {
		const { x, y, rect, kind } = this.#drag!;
		if (kind === 'line') {
			for (const [name, value] of Object.entries({ x1: x, y1: y, x2: point.x, y2: point.y }))
				rect.setAttribute(name, String(value));
			return;
		}
		rect.setAttribute('x', String(Math.min(x, point.x)));
		rect.setAttribute('y', String(Math.min(y, point.y)));
		rect.setAttribute('width', String(Math.abs(point.x - x)));
		rect.setAttribute('height', String(Math.abs(point.y - y)));
	}
	#cancel(): void {
		const drag = this.#drag;
		drag?.rect.remove();
		this.#drag = undefined;
		if (drag && this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const end = pagePoint(drag.svg, drag.page, event) ?? drag;
		this.#cancel();
		const width = Math.abs(end.x - drag.x),
			height = Math.abs(end.y - drag.y);
		if (
			drag.kind === 'line'
				? Math.hypot(width, height) < MIN_SIZE
				: width < MIN_SIZE || height < MIN_SIZE
		) {
			this.options.announce(`Drag on the page to draw a ${drag.kind}.`);
			return;
		}
		const state = this.controller.state;
		if (state.document?.pages[state.pageIndex] !== drag.page || this.options.tool() !== drag.kind)
			return;
		const request = ++this.#request;
		try {
			const shapeId =
				drag.kind === 'line'
					? await insertLine(this.controller, drag.page, drag, end)
					: await insertRectangle(
							this.controller,
							drag.page,
							{ x: (drag.x + end.x) / 2, y: (drag.y + end.y) / 2 },
							{ width, height },
						);
			if (request !== this.#request) return;
			this.options.announce(
				drag.kind === 'line'
					? `Line ${shapeId} added.`
					: `Rectangle ${shapeId} added (${+width.toFixed(4)} × ${+height.toFixed(4)} in).`,
			);
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		}
	}
}
