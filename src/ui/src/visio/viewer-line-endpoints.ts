import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	visioPageEditToDrawing,
	visioStraightLineHandles,
	visioSelectionIsOnPage,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { pagePoint } from './viewer-draw-tool';
import { handleGestureIsCurrent, wireHandleEvents } from './viewer-handle-events';

const SVG = 'http://www.w3.org/2000/svg';
const pointOptions = { snap: false, bounded: false } as const;

/** A preview-only endpoint gesture; release commits one source-backed core transaction. */
export class ViewerLineEndpoints {
	#drag:
		| {
				pointer: number;
				svg: SVGSVGElement;
				page: VisioPage;
				document: VisioDocument;
				shapeId: string;
				endpoint: 'begin' | 'end';
				preview: SVGLineElement;
				startX: number;
				startY: number;
		  }
		| undefined;
	#request = 0;
	#overlay: SVGGElement | undefined;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; announce(message: string): void },
	) {}
	render(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		if (this.#drag) {
			if (handleGestureIsCurrent(this.#drag, state, this.options.active())) return;
			this.#cancel();
		}
		this.#overlay?.remove();
		this.#overlay = undefined;
		if (
			!page ||
			!state.selectedShape ||
			!visioSelectionIsOnPage(state.selectedShape, page.id) ||
			state.selectedShapes.length !== 1 ||
			!state.edit.sourceAvailable ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const shape = page.shapes.find((shape) => shape.id === state.selectedShape!.id);
		if (
			!shape ||
			!visioStraightLineHandles(shape) ||
			page.connectors.some(
				(connection) => connection.fromShapeId === shape.id || connection.toShapeId === shape.id,
			)
		)
			return;
		const group = Array.from(this.viewport.querySelectorAll<SVGGElement>('[data-shape-id]')).find(
			(group) =>
				group.dataset.shapeId === shape.id &&
				(group.dataset.pageId ?? page.id) === page.id &&
				group.dataset.selected === 'true',
		);
		if (!group) return;
		const svg = group.ownerSVGElement;
		if (!svg) return;
		const overlay = this.viewport.ownerDocument.createElementNS(SVG, 'g');
		overlay.dataset.lineEndpointOverlay = '';
		const rootMatrix = svg.getCTM?.(),
			groupMatrix = group.getCTM?.();
		if (rootMatrix && groupMatrix) {
			const matrix = rootMatrix.inverse().multiply(groupMatrix);
			overlay.setAttribute(
				'transform',
				`matrix(${matrix.a} ${matrix.b} ${matrix.c} ${matrix.d} ${matrix.e} ${matrix.f})`,
			);
		}
		const screen = group.getScreenCTM?.();
		const screenScale = screen ? Math.hypot(screen.a, screen.b) : 96 * state.zoom;
		for (const endpoint of ['begin', 'end'] as const) {
			const handle = this.viewport.ownerDocument.createElementNS(SVG, 'circle');
			handle.dataset.lineEndpoint = endpoint;
			handle.dataset.lineShapeId = shape.id;
			handle.setAttribute('cx', String(endpoint === 'begin' ? 0 : shape.width));
			handle.setAttribute('cy', '0');
			handle.setAttribute('r', String(5 / screenScale));
			handle.setAttribute('vector-effect', 'non-scaling-stroke');
			handle.setAttribute('aria-label', `Drag ${endpoint} endpoint`);
			overlay.append(handle);
		}
		svg.append(overlay);
		this.#overlay = overlay;
	}
	wire(): () => void {
		const dispose = wireHandleEvents(this.viewport, {
			pointer: () => this.#drag?.pointer,
			start: (event) => this.#start(event),
			move: (event) => this.#move(event),
			finish: (event) => this.#finish(event),
			cancel: () => this.#cancel(),
		});
		return () => {
			++this.#request;
			dispose();
			this.#overlay?.remove();
			this.#overlay = undefined;
		};
	}

	#start(event: PointerEvent): void {
		const handle = (event.target as Element)?.closest?.<SVGCircleElement>('[data-line-endpoint]');
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		const endpoint = handle?.dataset.lineEndpoint;
		if (
			!handle ||
			!svg ||
			!page ||
			!state.document ||
			!state.selectedShape ||
			!visioSelectionIsOnPage(state.selectedShape, page.id) ||
			handle.dataset.lineShapeId !== state.selectedShape.id ||
			state.selectedShapes.length !== 1 ||
			this.#drag ||
			event.button !== 0 ||
			state.edit.busy ||
			state.loading ||
			!state.edit.sourceAvailable ||
			!this.options.active() ||
			(endpoint !== 'begin' && endpoint !== 'end')
		)
			return;
		const other = handle.parentElement?.querySelector<SVGCircleElement>(
			`[data-line-endpoint="${endpoint === 'begin' ? 'end' : 'begin'}"]`,
		);
		const matrix = other?.getScreenCTM();
		if (!other || !matrix) return;
		const centre = new DOMPoint(other.cx.baseVal.value, 0).matrixTransform(matrix);
		const fixed = pagePoint(svg, page, { clientX: centre.x, clientY: centre.y }, pointOptions);
		const start = pagePoint(svg, page, event, pointOptions);
		if (!fixed || !start) return;
		event.preventDefault();
		event.stopImmediatePropagation();
		const preview = this.viewport.ownerDocument.createElementNS(SVG, 'line');
		preview.classList.add('endpoint-preview');
		preview.setAttribute('vector-effect', 'non-scaling-stroke');
		for (const [name, value] of Object.entries({
			x1: fixed.x,
			y1: fixed.y,
			x2: start.x,
			y2: start.y,
		}))
			preview.setAttribute(name, String(value));
		svg.append(preview);
		this.#drag = {
			pointer: event.pointerId,
			svg,
			page,
			document: state.document,
			shapeId: state.selectedShape.id,
			endpoint,
			preview,
			startX: event.clientX,
			startY: event.clientY,
		};
		this.viewport.setPointerCapture?.(event.pointerId);
	}
	#move(event: PointerEvent): void {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!handleGestureIsCurrent(drag, this.controller.state, this.options.active())) {
			this.#cancel();
			return;
		}
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		if (point) {
			drag.preview.setAttribute('x2', String(point.x));
			drag.preview.setAttribute('y2', String(point.y));
		}
		event.preventDefault();
		event.stopImmediatePropagation();
	}
	#cancel(): void {
		const drag = this.#drag;
		this.#drag = undefined;
		drag?.preview.remove();
		if (drag && this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!handleGestureIsCurrent(drag, this.controller.state, this.options.active())) {
			this.#cancel();
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		this.#cancel();
		if (
			!point ||
			Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 0.5 ||
			this.controller.state.document !== drag.document
		)
			return;
		const request = ++this.#request;
		try {
			await this.controller.applyEdits([
				visioPageEditToDrawing(drag.page, {
					type: 'move-line-endpoint',
					pageId: drag.page.id,
					shapeId: drag.shapeId,
					endpoint: drag.endpoint,
					x: point.x,
					y: drag.page.height - point.y,
				}),
			]);
			if (request === this.#request)
				this.options.announce(`${drag.endpoint === 'begin' ? 'Begin' : 'End'} endpoint moved.`);
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		}
	}
}
