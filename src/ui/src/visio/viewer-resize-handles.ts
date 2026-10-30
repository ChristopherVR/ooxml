import { canEditDiagram } from './controller';
import { swimlaneResizeEdits } from './viewer-swimlanes';
import type { VisioDocument, VisioPage } from 'ooxml-core/visio';
import {
	visioResizeSelectionFrame,
	visioResizeSelectionDrag,
	visioSnapResizeDrag,
	visioSelectionIsOnPage,
	editErrorMessage,
	isEditCancellation,
	type VisioResizeHandle,
	type VisioPagePoint,
	type VisioResizeFrame,
	type VisioResizeSizeMatch,
	type VisioResizeSnapOptions,
	type VisioSnapResizeResult,
	type VisioSnapResult,
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
	shapeIds: string[];
	original: VisioResizeFrame;
	handle: VisioResizeHandle;
	selection: ViewerState['selectedShapes'];
	layers: ViewerState['layerVisibilityOverrides'];
	zoom: number;
	start: VisioPagePoint;
	frame: ReturnType<typeof createResizeFrame>;
	showOverlays?: () => void;
}
const pointOptions = { snap: false, bounded: false } as const;
/** Eight source-backed selection handles; saved content is never speculatively resized.
 * Like Visio, corners keep the proportions and Shift keeps them on every handle. */
export class ViewerResizeHandles {
	#drag: Drag | undefined;
	#frame: ReturnType<typeof createResizeFrame> | undefined;
	#request = 0;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: {
			active(): boolean;
			announce(message: string): void;
			/**
			 * Dynamic Grid and Guides for a resize: `settings` is undefined when snapping is off,
			 * `hint` draws the alignment lines and size matches, `clear` removes them.
			 */
			snap?: {
				settings(): VisioResizeSnapOptions | undefined;
				hint(
					page: VisioPage,
					lines: VisioSnapResult['lines'],
					matches: VisioResizeSizeMatch[],
				): void;
				clear(): void;
			};
		},
	) {}
	/**
	 * The resize solver for a pointer position. Ctrl resizes about the centre and Shift keeps the
	 * proportions; holding Alt leaves snapping out for this drag, as it does for a move.
	 */
	#solve(
		drag: Drag,
		point: VisioPagePoint,
		event: PointerEvent,
	): VisioSnapResizeResult | undefined {
		const snap = event.altKey ? undefined : this.options.snap?.settings();
		return visioSnapResizeDrag(drag.page, drag.shapeIds, drag.handle, drag.start, point, {
			proportional: event.shiftKey,
			symmetric: event.ctrlKey,
			...(snap ? { snap } : {}),
		});
	}
	#current(drag: Drag, state = this.controller.state): boolean {
		return (
			state.document === drag.document &&
			state.document.pages[state.pageIndex] === drag.page &&
			state.selectedShapes === drag.selection &&
			state.layerVisibilityOverrides === drag.layers &&
			state.zoom === drag.zoom &&
			canEditDiagram(state) &&
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
			!state.selectedShapes.every((item) => visioSelectionIsOnPage(item, page.id)) ||
			!canEditDiagram(state) ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const frame = visioResizeSelectionFrame(
			page,
			state.selectedShapes.map((item) => item.id),
		);
		if (!frame) return;
		const groups = Array.from(this.viewport.querySelectorAll<SVGGElement>('[data-shape-id]'));
		const selected = state.selectedShapes.map((item) =>
			groups.find(
				(group) =>
					group.dataset.shapeId === item.id &&
					(group.dataset.pageId ?? page.id) === page.id &&
					group.dataset.selected === 'true',
			),
		);
		const group = selected[0];
		const svg = group?.ownerSVGElement,
			screen = svg?.getScreenCTM?.();
		if (!svg || selected.some((group) => group?.ownerSVGElement !== svg)) return;
		const scale = screen ? Math.hypot(screen.a, screen.b) : 96 * state.zoom;
		if (!(scale > 0 && Number.isFinite(scale))) return;
		this.#frame = createResizeFrame(svg, page, frame, scale, selection.id);
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
			!state.selectedShapes.every((item) => visioSelectionIsOnPage(item, page.id)) ||
			event.button !== 0 ||
			this.#drag ||
			handle.dataset.resizeShape !== selection.id ||
			!canEditDiagram(state) ||
			state.edit.busy ||
			state.loading ||
			!this.options.active()
		)
			return;
		const start = pagePoint(svg, page, event, pointOptions),
			id = handle.dataset.resizeHandle as VisioResizeHandle;
		const shapeIds = state.selectedShapes.map((item) => item.id);
		const original = visioResizeSelectionFrame(page, shapeIds);
		if (!start || !original || !visioResizeSelectionDrag(page, shapeIds, id, start, start)) return;
		const drag: Drag = {
			pointer: event.pointerId,
			svg,
			document: state.document,
			page,
			shapeIds,
			original,
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
		const resized = point && this.#solve(drag, point, event);
		if (resized && this.#current(drag)) {
			drag.frame.update(resized.frame);
			this.options.snap?.hint(drag.page, resized.lines, resized.sizeMatches);
		}
		event.preventDefault();
		event.stopImmediatePropagation();
	}
	#cancel(): void {
		this.#release()();
	}
	/**
	 * Ends the gesture and returns the removal of its previews, which a finished drag keeps until its
	 * source edit settles (removing them first showed the old geometry for as long as the edit took).
	 */
	#release(): () => void {
		const drag = this.#drag;
		this.#drag = undefined;
		this.options.snap?.clear();
		if (!drag) return () => {};
		if (this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
		// Detached, so the renders while the edit runs do not remove the preview early.
		if (this.#frame === drag.frame) this.#frame = undefined;
		return () => {
			drag.showOverlays?.();
			drag.frame.dispose();
		};
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		const resized = point && this.#current(drag) && this.#solve(drag, point, event);
		event.preventDefault();
		event.stopImmediatePropagation();
		const settle = this.#release();
		if (!resized || !this.#current(drag)) {
			settle();
			this.render(this.controller.state);
			return;
		}
		if (
			Math.abs(resized.frame.width - drag.original.width) < 1e-10 &&
			Math.abs(resized.frame.height - drag.original.height) < 1e-10
		) {
			settle();
			this.render(this.controller.state);
			return;
		}
		// A lane resizes with the lanes after it moving and the frame following.
		const lane = swimlaneResizeEdits(drag.page, drag.shapeIds, resized.frame);
		if (lane === undefined) {
			settle();
			this.options.announce('Resize a swimlane by its own handles; the frame follows its lanes.');
			this.render(this.controller.state);
			return;
		}
		const request = ++this.#request;
		try {
			if (lane) await this.controller.applyEdits(lane);
			else
				await this.controller.applySelectionEdits(resized.commands, {
					pageId: drag.page.id,
					members: resized.members,
				});
			if (request === this.#request) this.options.announce('Shape resized.');
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		} finally {
			settle();
		}
	}
}
