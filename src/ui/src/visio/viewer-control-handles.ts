import { canEditDiagram } from './controller';
import type { VisioControl, VisioDocument, VisioPage, VisioShape } from 'ooxml-core/visio';
import { editErrorMessage, isEditCancellation, visioSelectionIsOnPage } from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { handleGestureIsCurrent, wireHandleEvents } from './viewer-handle-events';

const SVG = 'http://www.w3.org/2000/svg';
interface ControlDrag {
	pointer: number;
	svg: SVGSVGElement;
	page: VisioPage;
	document: VisioDocument;
	shapeId: string;
	control: VisioControl;
	handle: SVGRectElement;
	overlay: SVGGElement;
	startX: number;
	startY: number;
	x: number;
	y: number;
}

/** The local point under a client position, in the overlay's (shape-local, y-up) space. */
function localPoint(
	overlay: SVGGElement,
	clientX: number,
	clientY: number,
): { x: number; y: number } | undefined {
	const matrix = overlay.getScreenCTM?.();
	if (!matrix || !Number.isFinite(matrix.a + matrix.d)) return undefined;
	const point = new DOMPoint(clientX, clientY).matrixTransform(matrix.inverse());
	return { x: point.x, y: point.y };
}

/**
 * Visio's yellow control handles of the selected shape (its Controls rows). Dragging one along
 * its free axis shows the handle at the pointer; release commits one `set-control` source edit,
 * which rewrites the row's X or Y and recomputes the geometry that follows it. A handle the core
 * cannot move (fixed on both axes, or driven by a formula) is drawn grey and says why.
 */
export class ViewerControlHandles {
	#drag: ControlDrag | undefined;
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
			!canEditDiagram(state) ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const shape: VisioShape | undefined = page.shapes.find(
			(item) => item.id === state.selectedShape!.id,
		);
		if (!shape?.controls?.length) return;
		const group = Array.from(this.viewport.querySelectorAll<SVGGElement>('[data-shape-id]')).find(
			(item) =>
				item.dataset.shapeId === shape.id &&
				(item.dataset.pageId ?? page.id) === page.id &&
				item.dataset.selected === 'true',
		);
		const svg = group?.ownerSVGElement;
		if (!group || !svg) return;
		const overlay = this.viewport.ownerDocument.createElementNS(SVG, 'g');
		overlay.dataset.controlOverlay = '';
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
		const scale = screen ? Math.hypot(screen.a, screen.b) : 96 * state.zoom;
		const half = 4 / scale;
		for (const control of shape.controls) {
			const handle = this.viewport.ownerDocument.createElementNS(SVG, 'rect');
			handle.dataset.controlHandle = control.row;
			handle.dataset.controlShapeId = shape.id;
			if (control.locked) handle.dataset.controlLocked = '';
			handle.setAttribute('x', String(control.x - half));
			handle.setAttribute('y', String(control.y - half));
			handle.setAttribute('width', String(2 * half));
			handle.setAttribute('height', String(2 * half));
			handle.setAttribute('vector-effect', 'non-scaling-stroke');
			handle.setAttribute(
				'aria-label',
				control.locked
					? `Control handle (cannot be moved): ${control.locked}`
					: `Control handle${control.prompt ? `: ${control.prompt}` : ''}`,
			);
			const title = this.viewport.ownerDocument.createElementNS(SVG, 'title');
			title.textContent = control.locked ?? control.prompt ?? 'Drag to change the shape';
			handle.append(title);
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
		const handle = (event.target as Element)?.closest?.<SVGRectElement>('[data-control-handle]');
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		const overlay = handle?.parentElement as SVGGElement | null | undefined;
		const shape = page?.shapes.find((item) => item.id === handle?.dataset.controlShapeId);
		const control = shape?.controls?.find((item) => item.row === handle?.dataset.controlHandle);
		if (
			!handle ||
			!svg ||
			!overlay ||
			!page ||
			!state.document ||
			!control ||
			!state.selectedShape ||
			!visioSelectionIsOnPage(state.selectedShape, page.id) ||
			handle.dataset.controlShapeId !== state.selectedShape.id ||
			state.selectedShapes.length !== 1 ||
			this.#drag ||
			event.button !== 0 ||
			state.edit.busy ||
			state.loading ||
			!canEditDiagram(state) ||
			!this.options.active()
		)
			return;
		event.preventDefault();
		event.stopImmediatePropagation();
		if (control.locked) {
			this.options.announce(`This control handle cannot be moved. ${control.locked}`);
			return;
		}
		this.#drag = {
			pointer: event.pointerId,
			svg,
			page,
			document: state.document,
			shapeId: state.selectedShape.id,
			control,
			handle,
			overlay,
			startX: event.clientX,
			startY: event.clientY,
			x: control.x,
			y: control.y,
		};
		this.viewport.setPointerCapture?.(event.pointerId);
		if (control.prompt) this.options.announce(control.prompt);
	}

	#move(event: PointerEvent): void {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!handleGestureIsCurrent(drag, this.controller.state, this.options.active())) {
			this.#cancel();
			return;
		}
		const point = localPoint(drag.overlay, event.clientX, event.clientY);
		if (point) {
			// Only the free axes follow the pointer, as in Visio.
			if (drag.control.drag.x) drag.x = point.x;
			if (drag.control.drag.y) drag.y = point.y;
			const half = drag.handle.width.baseVal.value / 2;
			drag.handle.setAttribute('x', String(drag.x - half));
			drag.handle.setAttribute('y', String(drag.y - half));
		}
		event.preventDefault();
		event.stopImmediatePropagation();
	}

	#cancel(): void {
		this.#release()();
	}

	/** Ends the gesture and returns the reset of the handle, which a finished drag keeps until its edit settles. */
	#release(): () => void {
		const drag = this.#drag;
		this.#drag = undefined;
		if (drag && this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
		return () => {
			if (!drag) return;
			const half = drag.handle.width.baseVal.value / 2;
			drag.handle.setAttribute('x', String(drag.control.x - half));
			drag.handle.setAttribute('y', String(drag.control.y - half));
		};
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
		const point = localPoint(drag.overlay, event.clientX, event.clientY);
		const settle = this.#release();
		if (
			!point ||
			Math.hypot(event.clientX - drag.startX, event.clientY - drag.startY) < 0.5 ||
			this.controller.state.document !== drag.document
		) {
			settle();
			return;
		}
		const request = ++this.#request;
		try {
			await this.controller.applyEdits([
				{
					type: 'set-control',
					pageId: drag.page.id,
					shapeId: drag.shapeId,
					row: drag.control.row,
					...(drag.control.drag.x ? { x: point.x } : {}),
					...(drag.control.drag.y ? { y: point.y } : {}),
				},
			]);
			if (request === this.#request) this.options.announce('Control handle moved.');
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		} finally {
			settle();
		}
	}
}
