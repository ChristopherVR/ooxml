import type { VisioPage, VisioDocument } from 'ooxml-core/visio';
import {
	visioDrawBounds,
	visioDrawIsLargeEnough,
	editErrorMessage,
	isEditCancellation,
	isVisioPathTool,
	visioPathPreview,
	type VisioDrawingPoint,
	type VisioDrawingTool,
	type VisioPathTool,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import type { ViewerCreationToken } from './creation-token';
import { pagePoint } from './viewer-page-point';
import { wireHandleEvents } from './viewer-handle-events';

export interface DrawingGesture {
	pointer: number;
	svg: SVGSVGElement;
	page: VisioPage;
	document: VisioDocument;
	start: VisioDrawingPoint;
	kind: VisioDrawingTool | VisioPathTool;
	/** Pointer samples of a Freeform, Pencil or Arc stroke, start first. */
	points: VisioDrawingPoint[];
	token: ViewerCreationToken;
	preview: SVGRectElement | SVGEllipseElement | SVGLineElement | SVGPathElement;
	sourceGeneration: number;
	selection: ViewerState['selectedShapes'];
	zoom: number;
	layers: ViewerState['layerVisibilityOverrides'];
}
const MAX_SAMPLES = 10_000;
/** One captured drawing intent. Core plans geometry; the DOM owns capture and a temporary frame. */
export class ViewerDrawingGesture {
	#drag: DrawingGesture | undefined;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: {
			tool(): VisioDrawingTool | VisioPathTool | undefined;
			finish(drag: DrawingGesture, end: VisioDrawingPoint): Promise<void>;
			existingShape?(event: PointerEvent): boolean;
			announce(message: string): void;
		},
	) {}
	get drawing(): boolean {
		return !!this.#drag;
	}
	current(drag: DrawingGesture): boolean {
		const state = this.controller.state;
		return (
			this.controller.isCreationTokenCurrent(drag.token) &&
			state.document === drag.document &&
			state.document?.pages[state.pageIndex] === drag.page &&
			state.selectedShapes === drag.selection &&
			state.zoom === drag.zoom &&
			state.layerVisibilityOverrides === drag.layers &&
			this.controller.sourceGeneration === drag.sourceGeneration &&
			drag.svg.isConnected &&
			this.options.tool() === drag.kind &&
			this.controller.isCreationTokenCurrent(drag.token)
		);
	}
	render(_state: ViewerState): void {
		if (this.#drag && !this.current(this.#drag)) this.cancel();
	}
	wire(): () => void {
		return wireHandleEvents(this.viewport, {
			pointer: () => this.#drag?.pointer,
			start: (event) => this.#start(event),
			move: (event) => this.#move(event),
			finish: (event) => this.#finish(event),
			cancel: () => this.cancel(),
			consumeClick: () => !!this.options.tool(),
		});
	}
	cancel(): void {
		const drag = this.#drag;
		this.#drag = undefined;
		drag?.preview.remove();
		if (drag && this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
	}
	#start(event: PointerEvent): void {
		const kind = this.options.tool();
		if (!kind || this.#drag || event.button !== 0) return;
		const state = this.controller.state;
		if (!state.edit.sourceAvailable) return;
		const sourceGeneration = this.controller.sourceGeneration;
		const page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		if (
			!page ||
			!state.document ||
			!svg ||
			state.loading ||
			state.edit.busy ||
			!state.edit.sourceAvailable
		)
			return;
		if (this.options.existingShape?.(event)) return;
		// Freehand samples stay unsnapped so the fitted curve follows the pointer.
		const freehand = kind === 'freeform' || kind === 'pencil';
		const start = pagePoint(svg, page, event, { bounded: false, snap: !freehand });
		const current = this.controller.state;
		if (
			!start ||
			current.document !== state.document ||
			!current.edit.sourceAvailable ||
			current.document?.pages[current.pageIndex] !== page ||
			current.selectedShapes !== state.selectedShapes ||
			current.zoom !== state.zoom ||
			current.layerVisibilityOverrides !== state.layerVisibilityOverrides
		)
			return;
		let token: ViewerCreationToken;
		try {
			token = this.controller.captureCreationToken(page.id);
		} catch (error) {
			if (!isEditCancellation(error)) this.options.announce(editErrorMessage(error));
			return;
		}
		if (
			this.controller.state.document !== state.document ||
			!this.controller.isCreationTokenCurrent(token)
		)
			return;
		const preview = this.viewport.ownerDocument.createElementNS(
			'http://www.w3.org/2000/svg',
			isVisioPathTool(kind)
				? 'path'
				: kind === 'line'
					? 'line'
					: kind === 'ellipse'
						? 'ellipse'
						: 'rect',
		);
		preview.classList.add('draw-preview');
		preview.setAttribute('vector-effect', 'non-scaling-stroke');
		const drag: DrawingGesture = {
			pointer: event.pointerId,
			svg,
			page,
			document: state.document,
			start,
			kind,
			points: [start],
			token,
			preview,
			sourceGeneration,
			selection: state.selectedShapes,
			zoom: state.zoom,
			layers: state.layerVisibilityOverrides,
		};
		if (!this.current(drag)) return;
		event.preventDefault();
		event.stopImmediatePropagation();
		svg.append(preview);
		this.#drag = drag;
		this.#update(start);
		this.viewport.focus({ preventScroll: true });
		if (this.current(drag)) this.viewport.setPointerCapture?.(event.pointerId);
		else this.cancel();
	}
	#move(event: PointerEvent): void {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!this.current(drag)) {
			this.cancel();
			return;
		}
		const point = this.#point(drag, event);
		if (point && this.current(drag)) this.#update(point);
	}
	#point(drag: DrawingGesture, event: PointerEvent): VisioDrawingPoint | undefined {
		const freehand = drag.kind === 'freeform' || drag.kind === 'pencil';
		return pagePoint(drag.svg, drag.page, event, { bounded: false, snap: !freehand });
	}
	#update(point: VisioDrawingPoint): void {
		const drag = this.#drag!;
		if (isVisioPathTool(drag.kind)) {
			if (drag.kind === 'arc') drag.points = [drag.start, point];
			else if (point !== drag.start && drag.points.length < MAX_SAMPLES) drag.points.push(point);
			drag.preview.setAttribute('d', visioPathPreview(drag.kind, drag.points));
			return;
		}
		const bounds = visioDrawBounds(drag.start, point);
		const attributes =
			drag.kind === 'line'
				? { x1: drag.start.x, y1: drag.start.y, x2: point.x, y2: point.y }
				: drag.kind === 'ellipse'
					? { cx: bounds.centerX, cy: bounds.centerY, rx: bounds.radiusX, ry: bounds.radiusY }
					: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height };
		for (const [name, value] of Object.entries(attributes))
			drag.preview.setAttribute(name, String(value));
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const end = this.#point(drag, event),
			current = this.current(drag);
		this.cancel();
		if (!current || !end || !this.current(drag)) return;
		if (isVisioPathTool(drag.kind)) {
			if (drag.kind === 'arc') drag.points = [drag.start, end];
			else if (drag.points.length < MAX_SAMPLES) drag.points.push(end);
		}
		try {
			if (!isVisioPathTool(drag.kind) && !visioDrawIsLargeEnough(drag.kind, drag.start, end)) {
				this.options.announce(
					`Drag on the page to draw a ${drag.kind === 'text' ? 'text box' : drag.kind}.`,
				);
				return;
			}
			if (this.current(drag)) await this.options.finish(drag, end);
		} catch (error) {
			if (this.current(drag) && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		}
	}
}
