import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import {
	editErrorMessage,
	isEditCancellation,
	visioPageEditToDrawing,
	visioStraightLineHandles,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { pagePoint } from './viewer-draw-tool';

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
	#suppressClick = false;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; announce(message: string): void },
	) {}
	render(state: ViewerState): void {
		const page = state.document?.pages[state.pageIndex];
		if (this.#drag) {
			if (
				state.document === this.#drag.document &&
				page === this.#drag.page &&
				state.selectedShape?.id === this.#drag.shapeId &&
				this.#drag.svg.isConnected &&
				!state.loading &&
				!state.edit.busy &&
				this.options.active()
			)
				return;
			this.#cancel();
		}
		for (const overlay of this.viewport.querySelectorAll('[data-line-endpoint-overlay]'))
			overlay.remove();
		if (
			!page ||
			!state.selectedShape ||
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
			(group) => group.dataset.shapeId === shape.id && group.dataset.selected === 'true',
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
	}
	wire(): () => void {
		const Abort = this.viewport.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal, capture: true };
		this.viewport.addEventListener('pointerdown', (event) => this.#start(event), options);
		this.viewport.addEventListener('pointermove', (event) => this.#move(event), options);
		this.viewport.addEventListener('pointerup', (event) => void this.#finish(event), options);
		this.viewport.addEventListener(
			'pointercancel',
			(event) => {
				if (event.pointerId === this.#drag?.pointer) this.#cancel();
			},
			options,
		);
		this.viewport.addEventListener(
			'lostpointercapture',
			(event) => {
				if (event.pointerId === this.#drag?.pointer) this.#cancel();
			},
			options,
		);
		this.viewport.addEventListener(
			'keydown',
			(event) => {
				if (event.key === 'Escape' && this.#drag) {
					event.preventDefault();
					event.stopImmediatePropagation();
					this.#cancel();
				}
			},
			options,
		);
		this.viewport.addEventListener(
			'click',
			(event) => {
				if (this.#suppressClick) {
					this.#suppressClick = false;
					event.stopImmediatePropagation();
				}
			},
			options,
		);
		return () => {
			++this.#request;
			this.#cancel();
			events.abort();
			for (const overlay of this.viewport.querySelectorAll('[data-line-endpoint-overlay]'))
				overlay.remove();
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
		event.preventDefault();
		event.stopImmediatePropagation();
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		this.#cancel();
		this.#suppressClick = true;
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
