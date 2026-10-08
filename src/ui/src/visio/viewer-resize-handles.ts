import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import {
	visioResizeShape,
	visioResizeDrag,
	visioSelectionIsOnPage,
	editErrorMessage,
	isEditCancellation,
	type VisioResizeHandle,
	type VisioPagePoint,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { pagePoint } from './viewer-draw-tool';
import { wireHandleEvents } from './viewer-handle-events';
import { createResizeFrame } from './viewer-resize-frame';
import { hideGestureOverlays } from './viewer-gesture-overlays';
interface Drag {
	pointer: number;
	svg: SVGSVGElement;
	document: VisioDocument;
	page: VisioPage;
	shapeId: string;
	handle: VisioResizeHandle;
	selection: ViewerState['selectedShapes'];
	layers: ViewerState['layerVisibilityOverrides'];
	zoom: number;
	start: VisioPagePoint;
	frame: ReturnType<typeof createResizeFrame>;
	showOverlays?: () => void;
}
const pointOptions = { snap: false, bounded: false } as const;
/** Eight source-backed single-selection handles; saved content is never speculatively resized. */
export class ViewerResizeHandles {
	#drag: Drag | undefined;
	#frame: ReturnType<typeof createResizeFrame> | undefined;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; announce(message: string): void },
	) {}
	#current(drag: Drag, state = this.controller.state): boolean {
		return (
			state.document === drag.document &&
			state.document.pages[state.pageIndex] === drag.page &&
			state.selectedShapes === drag.selection &&
			state.layerVisibilityOverrides === drag.layers &&
			state.zoom === drag.zoom &&
			state.edit.sourceAvailable &&
			!state.edit.busy &&
			!state.loading &&
			drag.svg.isConnected &&
			this.options.active()
		);
	}
	render(state: ViewerState): void {
		if (this.#drag) {
			if (this.#current(this.#drag, state)) return;
			this.#cancel();
		}
		this.#frame?.dispose();
		this.#frame = undefined;
		const page = state.document?.pages[state.pageIndex],
			selection = state.selectedShape;
		if (
			!page ||
			!selection ||
			state.selectedShapes.length !== 1 ||
			!visioSelectionIsOnPage(selection, page.id) ||
			!state.edit.sourceAvailable ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const shape = visioResizeShape(page, selection.id);
		if (!shape) return;
		const group = Array.from(this.viewport.querySelectorAll<SVGGElement>('[data-shape-id]')).find(
			(group) =>
				group.dataset.shapeId === shape.id &&
				(group.dataset.pageId ?? page.id) === page.id &&
				group.dataset.selected === 'true',
		);
		const svg = group?.ownerSVGElement,
			screen = svg?.getScreenCTM?.();
		if (!svg) return;
		const scale = screen ? Math.hypot(screen.a, screen.b) : 96 * state.zoom;
		if (!(scale > 0 && Number.isFinite(scale))) return;
		this.#frame = createResizeFrame(svg, page, shape, scale);
	}
	wire(): () => void {
		const dispose = wireHandleEvents(this.viewport, {
			pointer: () => this.#drag?.pointer,
			start: (event) => this.#start(event),
			move: (event) => this.#move(event),
			finish: (event) => this.#finish(event),
			cancel: () => {
				this.#cancel();
				this.render(this.controller.state);
			},
		});
		return () => {
			++this.#request;
			dispose();
			this.#frame?.dispose();
			this.#frame = undefined;
		};
	}
	#start(event: PointerEvent): void {
		const handle = (event.target as Element)?.closest<SVGCircleElement>('[data-resize-handle]');
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex],
			selection = state.selectedShape;
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		if (
			!handle ||
			!this.#frame?.element.contains(handle) ||
			!svg ||
			!page ||
			!state.document ||
			!selection ||
			!visioSelectionIsOnPage(selection, page.id) ||
			state.selectedShapes.length !== 1 ||
			event.button !== 0 ||
			this.#drag ||
			handle.dataset.resizeShape !== selection.id ||
			!state.edit.sourceAvailable ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const start = pagePoint(svg, page, event, pointOptions),
			id = handle.dataset.resizeHandle as VisioResizeHandle;
		if (!start || !visioResizeDrag(page, selection.id, id, start, start)) return;
		const drag: Drag = {
			pointer: event.pointerId,
			svg,
			document: state.document,
			page,
			shapeId: selection.id,
			handle: id,
			selection: state.selectedShapes,
			layers: state.layerVisibilityOverrides,
			zoom: state.zoom,
			start,
			frame: this.#frame,
		};
		if (!this.#current(drag)) return;
		event.preventDefault();
		event.stopImmediatePropagation();
		this.#drag = drag;
		drag.showOverlays = hideGestureOverlays(this.viewport, '.rotation-overlay');
		this.#frame.rectangle.dataset.resizeFramePreview = '';
		this.viewport.focus({ preventScroll: true });
		if (this.#drag === drag && this.#current(drag))
			this.viewport.setPointerCapture?.(event.pointerId);
	}
	#move(event: PointerEvent): void {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!this.#current(drag)) {
			this.#cancel();
			return;
		}
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		const resized =
			point && visioResizeDrag(drag.page, drag.shapeId, drag.handle, drag.start, point);
		if (resized && this.#current(drag)) drag.frame.update(resized.frame);
		event.preventDefault();
		event.stopImmediatePropagation();
	}
	#cancel(): void {
		const drag = this.#drag;
		this.#drag = undefined;
		if (drag) {
			drag.showOverlays?.();
			drag.frame.dispose();
			if (this.#frame === drag.frame) this.#frame = undefined;
			if (this.viewport.hasPointerCapture?.(drag.pointer))
				this.viewport.releasePointerCapture(drag.pointer);
		}
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		const resized =
			point &&
			this.#current(drag) &&
			visioResizeDrag(drag.page, drag.shapeId, drag.handle, drag.start, point);
		event.preventDefault();
		event.stopImmediatePropagation();
		this.#cancel();
		if (!resized || !this.#current(drag)) {
			this.render(this.controller.state);
			return;
		}
		const original = drag.page.shapes.find((shape) => shape.id === drag.shapeId)!;
		if (
			Math.abs(resized.frame.width - original.width) < 1e-10 &&
			Math.abs(resized.frame.height - original.height) < 1e-10
		) {
			this.render(this.controller.state);
			return;
		}
		const request = ++this.#request;
		try {
			await this.controller.applySelectionEdits([resized.command]);
			if (request === this.#request) this.options.announce('Shape resized.');
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		}
	}
}
