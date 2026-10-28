import type { VisioPage } from 'ooxml-core/visio';
import {
	visioMoveCommands,
	visioMoveTargets,
	visioMovementShape,
	visioPageDragDelta,
	visioMarqueeBox,
	visioSelectionIsOnPage,
	editErrorMessage,
	isEditCancellation,
	type VisioPagePoint,
} from 'ooxml-core/visio/ui';
import type { ViewerController, ViewerState } from './controller';
import { pagePoint } from './viewer-draw-tool';
import { wireHandleEvents } from './viewer-handle-events';
import { createMovementPreview } from './viewer-movement-preview';
import { hideGestureOverlays } from './viewer-gesture-overlays';
import { pointerShapeTarget } from './viewer-shape-target';
import { trackPointerPan } from '../pointer-gesture';
import { ViewerTouchGestures } from './viewer-touch-gestures';
import { TOUCH_SLOP } from './viewer-touch-math';
import { armLongPress, touchRelease } from './viewer-touch-long-press';
import {
	createGesture,
	createMarqueeRect,
	marqueeSelect,
	gestureIsCurrent,
	selectDragTarget,
	type Gesture,
} from './viewer-gesture-state';
import { nearbyShape } from './viewer-touch-hit';

const pointOptions = { snap: false, bounded: false } as const;
/** Direct pointer movement and full-enclosure marquee; release is the only source edit. */
export class ViewerPointerGestures {
	#drag: Gesture | undefined;
	#request = 0;
	#cancelClick = false;
	#cancelPan: (() => void) | undefined;
	readonly #touch: ViewerTouchGestures;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: {
			active(): boolean;
			announce(message: string): void;
			/** Guides and Dynamic Grid: adjust a drag delta and show alignment hints. */
			snap?(
				page: VisioPage,
				ids: readonly string[],
				delta: VisioPagePoint,
				axis?: 'x' | 'y',
			): VisioPagePoint;
			clearSnap?(): void;
		},
	) {
		this.#touch = new ViewerTouchGestures(viewport, controller, {
			active: () => options.active(),
			cancelSingle: () => {
				this.#cancelPan?.();
				this.#cancelPan = undefined;
				if (this.#drag) this.#cancel();
				this.#cancelClick = true;
			},
		});
	}
	render(state: ViewerState): void {
		this.#cancelPan?.();
		if (this.#drag && !this.#current(this.#drag, state)) this.#cancel();
	}
	wire(): () => void {
		const stopOutside = (event: PointerEvent) => {
			if (this.#drag && !this.#drag.started && !event.composedPath().includes(this.viewport))
				this.#cancel();
		};
		this.viewport.ownerDocument.addEventListener('pointerup', stopOutside);
		// First, so a second finger reaches the pinch before the one-pointer gestures see it.
		const disposeTouch = this.#touch.wire();
		const dispose = wireHandleEvents(this.viewport, {
			pointer: () => this.#drag?.pointer,
			start: (event) => this.#start(event),
			move: (event) => this.#move(event),
			finish: (event) => this.#finish(event),
			cancel: () => this.#cancel(),
			shouldSuppressClick: () => !!this.#drag?.started,
			consumeClick: () => {
				const value = this.#cancelClick;
				this.#cancelClick = false;
				return value;
			},
		});
		return () => {
			this.#cancelPan?.();
			++this.#request;
			this.viewport.ownerDocument.removeEventListener('pointerup', stopOutside);
			dispose();
			disposeTouch();
		};
	}
	#current(drag: Gesture, state = this.controller.state): boolean {
		return gestureIsCurrent(drag, state, this.options.active());
	}
	#start(event: PointerEvent): void {
		this.#cancelClick = false;
		if (this.#drag || event.button !== 0 || !this.options.active()) return;
		const node = event.target as Element;
		if (
			node.closest?.(
				'[data-line-endpoint], [data-rotation-handle], [data-resize-handle], .rotation-overlay, [data-guide-id]',
			)
		)
			return;
		const state = this.controller.state,
			page = state.document?.pages[state.pageIndex];
		const svg = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		if (
			!svg ||
			!page ||
			!state.document ||
			state.loading ||
			state.edit.busy ||
			(!svg.contains(node) && !(node === this.viewport && event.pointerType === 'touch'))
		)
			return;
		// Dragging inside a group moves the whole top-level group, as in Visio.
		let group = pointerShapeTarget(node, state.selectedShapes, true);
		const touch = event.pointerType === 'touch';
		let nearby = false;
		if (!group && touch) {
			group = nearbyShape(this.viewport, svg, event, state.selectedShapes);
			nearby = !!group;
		}
		// A finger on empty paper pans; shape bodies keep their editing gesture.
		if (!group && touch) {
			this.#cancelPan?.();
			this.#cancelPan = trackPointerPan(
				this.viewport,
				event,
				() => {
					this.#cancelClick = true;
				},
				() => {
					this.#cancelPan = undefined;
				},
			);
			event.preventDefault();
			return;
		}
		if (
			group &&
			(!state.edit.sourceAvailable ||
				!state.canEdit ||
				(group.dataset.pageId ?? page.id) !== page.id ||
				!visioMovementShape(page, group.dataset.shapeId!))
		)
			return;
		const start = pagePoint(svg, page, event, pointOptions);
		if (!start) return;
		const drag = createGesture({ event, svg, page, state, start, touch, nearby, group });
		if (!this.#current(drag)) return;
		this.#drag = drag;
		if (touch && group)
			armLongPress(this.viewport.ownerDocument.defaultView, drag, group, event, () => {
				return this.#drag === drag && this.#current(drag);
			});
	}
	#begin(drag: Gesture): boolean {
		if (drag.target) {
			const selected = drag.selection.some(
				(shape) => shape.id === drag.target!.id && visioSelectionIsOnPage(shape, drag.page.id),
			);
			if (!selected) {
				this.#drag = undefined;
				if (!selectDragTarget(this.controller, drag, this.options.active())) return false;
				this.#drag = drag;
			}
			if (
				!drag.selection.every((shape) => visioSelectionIsOnPage(shape, drag.page.id)) ||
				!visioMoveCommands(
					drag.page,
					drag.selection.map((shape) => shape.id),
					{ x: 0, y: 0 },
				)
			) {
				this.options.announce(
					'The entire selection must contain movable local two-dimensional shapes.',
				);
				return false;
			}
			// Container members move with their container, so they preview with it too.
			drag.preview = createMovementPreview(
				drag.svg,
				drag.page,
				visioMoveTargets(
					drag.page,
					drag.selection.map((shape) => shape.id),
				),
			);
			drag.showOverlays = hideGestureOverlays(
				this.viewport,
				'[data-resize-overlay], .rotation-overlay',
			);
		} else {
			drag.marquee = createMarqueeRect(drag.svg);
		}
		drag.started = true;
		drag.held = undefined;
		this.viewport.focus({ preventScroll: true });
		this.viewport.setPointerCapture?.(drag.pointer);
		return true;
	}
	#move(event: PointerEvent): void {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		if (!this.#current(drag)) {
			this.#cancel();
			return;
		}
		if (!drag.started) {
			const slop = drag.touch ? TOUCH_SLOP : 4;
			if (Math.hypot(event.clientX - drag.clientX, event.clientY - drag.clientY) < slop) return;
			try {
				if (!this.#begin(drag)) {
					this.#cancel();
					return;
				}
			} catch (error) {
				this.#cancel();
				this.options.announce(editErrorMessage(error));
				return;
			}
		}
		if (this.#drag !== drag || !this.#current(drag)) {
			this.#cancel();
			return;
		}
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		if (point && this.#current(drag)) {
			if (drag.preview) drag.preview.update(this.#delta(drag, point, event));
			const box = drag.marquee && visioMarqueeBox(drag.start, point);
			if (box)
				for (const [name, value] of Object.entries(box))
					drag.marquee!.setAttribute(name, String(value));
		}
		event.preventDefault();
		event.stopImmediatePropagation();
	}
	#delta(drag: Gesture, point: VisioPagePoint, event: PointerEvent): VisioPagePoint {
		const delta = visioPageDragDelta(drag.start, point, event.shiftKey);
		if (!Number.isFinite(delta.x) || !Number.isFinite(delta.y)) return delta;
		if (event.altKey) {
			this.options.clearSnap?.();
			return delta;
		}
		const axis = event.shiftKey ? (delta.y === 0 ? 'x' : 'y') : undefined;
		const ids = drag.selection.map((shape) => shape.id);
		const snapped = this.options.snap?.(drag.page, ids, delta, axis) ?? delta;
		// A host-provided snap callback must not move the constrained axis either.
		return { x: axis === 'y' ? 0 : snapped.x, y: axis === 'x' ? 0 : snapped.y };
	}
	#cancel(): void {
		this.#release()();
	}
	/**
	 * Ends the gesture and returns the removal of its previews. A finished move keeps showing the
	 * preview until its source edit settles; removing it first showed the shape back at its old
	 * position for as long as the edit took.
	 */
	#release(): () => void {
		const drag = this.#drag;
		this.#drag = undefined;
		if (drag?.longPress) this.viewport.ownerDocument.defaultView?.clearTimeout(drag.longPress);
		if (drag?.started) this.#cancelClick = true;
		if (drag && this.viewport.hasPointerCapture?.(drag.pointer))
			this.viewport.releasePointerCapture(drag.pointer);
		return () => {
			drag?.preview?.dispose();
			drag?.showOverlays?.();
			drag?.marquee?.remove();
			if (drag?.preview) this.options.clearSnap?.();
		};
	}
	async #finish(event: PointerEvent): Promise<void> {
		const drag = this.#drag;
		if (!drag || event.pointerId !== drag.pointer) return;
		const tap =
			!drag.started && this.#current(drag) ? touchRelease(drag, this.controller) : undefined;
		if (tap) {
			this.#cancel();
			this.#cancelClick = true;
			tap(this.viewport.ownerDocument.defaultView);
			return;
		}
		if (!this.#current(drag) || !drag.started) {
			this.#cancel();
			return;
		}
		const point = pagePoint(drag.svg, drag.page, event, pointOptions);
		event.preventDefault();
		event.stopImmediatePropagation();
		const settle = this.#release();
		if (!point || !this.#current(drag) || !drag.target) settle();
		if (!point || !this.#current(drag)) return;
		if (!drag.target) {
			if (this.#current(drag)) marqueeSelect(this.controller, drag, point);
			return;
		}
		const edits = visioMoveCommands(
			drag.page,
			drag.selection.map((shape) => shape.id),
			this.#delta(drag, point, event),
		);
		this.options.clearSnap?.();
		if (!edits?.length || !this.#current(drag)) {
			settle();
			return;
		}
		const request = ++this.#request;
		try {
			await this.controller.applySelectionEdits(edits);
			if (request === this.#request) this.options.announce('Shapes moved.');
		} catch (error) {
			if (request === this.#request && !isEditCancellation(error))
				this.options.announce(editErrorMessage(error));
		} finally {
			settle();
		}
	}
}
