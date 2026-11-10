/**
 * Zooming and panning the slide viewport with a mouse or trackpad.
 *
 * Before this module, Ctrl+wheel (and a trackpad pinch, which browsers report
 * as Ctrl+wheel) zoomed only in React, and there it changed the zoom factor
 * without moving the scroll position: the slide grew from its top-left corner,
 * so whatever the pointer was over slid away. The other four bindings left the
 * gesture to the browser, which zoomed the whole page. And a zoomed slide could
 * only be moved with the scrollbars or a two-axis trackpad scroll: there was no
 * way to drag it, so a mouse could not reach the right-hand side of a slide.
 *
 * The behaviour is shared so the five bindings cannot drift: each binding
 * calls {@link attachViewportNavigation} with its `[data-pptx-viewport]`
 * element and its zoom store, and the returned function when the viewport
 * goes away.
 *
 * - **Zoom at the pointer.** Ctrl/Cmd+wheel, and Safari's pinch gesture
 *   events, zoom so that the point of the slide under the pointer stays under
 *   the pointer, like PowerPoint's Ctrl+wheel.
 * - **Pan by dragging.** Middle-button drag, or primary-button drag while
 *   Space is held, scrolls a slide that is larger than the viewport. When the
 *   binding reports that a plain drag has nothing to select (a read-only view),
 *   a plain drag pans too, once it has moved a few pixels so a click still
 *   reaches the slide.
 *
 * @module render/viewport-navigation
 */
import { createWheelStepBuffer, mapEditingWheel } from './wheel-intent';

/** The slide stage inside a viewport: the region every binding marks as the slide. */
export const VIEWPORT_STAGE_SELECTOR = '[aria-roledescription="slide"]';

/** Pointer travel, in CSS pixels, before a plain drag in a read-only view starts panning. */
export const DRAG_PAN_THRESHOLD_PX = 4;

/**
 * Frames to wait for the binding to lay the slide out at a new zoom before the
 * scroll correction is dropped. Every binding commits a store change within a
 * frame or two; this only bounds a zoom that did not change the slide's size.
 */
const MAX_LAYOUT_FRAMES = 6;

/**
 * A zoom step within this long of the last one, with the pointer still within
 * {@link ANCHOR_REUSE_DISTANCE_PX}, continues the same gesture and keeps its
 * slide point. While the slide is smaller than the viewport it stays centred,
 * so the point under the pointer drifts and no scroll can hold it; keeping the
 * gesture's first point brings that point back under the pointer as soon as the
 * slide is large enough to scroll.
 */
const ANCHOR_REUSE_MS = 400;
const ANCHOR_REUSE_DISTANCE_PX = 2;

/** The part of a `DOMRect` the anchoring maths reads. */
export interface ViewportRect {
	left: number;
	top: number;
	width: number;
	height: number;
}

/** A point of the slide pinned to a point of the screen while the zoom changes. */
export interface ZoomAnchor {
	/** Horizontal position on the slide, as a fraction of its width. */
	fx: number;
	/** Vertical position on the slide, as a fraction of its height. */
	fy: number;
	/** Client X the slide point should end up under. */
	clientX: number;
	/** Client Y the slide point should end up under. */
	clientY: number;
}

/**
 * Where a client point is on the slide, as fractions of the slide's size.
 * Null while the slide has no size (not laid out yet).
 */
export function zoomAnchorAt(
	stage: ViewportRect,
	clientX: number,
	clientY: number,
): ZoomAnchor | null {
	if (!(stage.width > 0) || !(stage.height > 0)) {
		return null;
	}
	return {
		fx: (clientX - stage.left) / stage.width,
		fy: (clientY - stage.top) / stage.height,
		clientX,
		clientY,
	};
}

/**
 * The viewport scroll offsets that put an anchored slide point back under its
 * client point, given where the slide is laid out now (after the zoom).
 *
 * The result can be negative or past the scroll range when the slide is
 * smaller than the viewport and centred in it; assigning it to `scrollLeft` /
 * `scrollTop` clamps it, which is the closest the viewport can get.
 */
export function anchoredScrollPosition(
	scroll: { left: number; top: number },
	stage: ViewportRect,
	anchor: ZoomAnchor,
): { left: number; top: number } {
	return {
		left: scroll.left + stage.left + anchor.fx * stage.width - anchor.clientX,
		top: scroll.top + stage.top + anchor.fy * stage.height - anchor.clientY,
	};
}

/** True when the viewport's content is larger than the viewport on either axis. */
export function viewportCanPan(
	viewport: Pick<HTMLElement, 'scrollWidth' | 'scrollHeight' | 'clientWidth' | 'clientHeight'>,
): boolean {
	return (
		viewport.scrollWidth > viewport.clientWidth || viewport.scrollHeight > viewport.clientHeight
	);
}

/**
 * Scroll a viewport to where a newly shown slide starts: its top-left corner.
 *
 * Every binding calls this when it shows another slide, and only then. The
 * zoom belongs to the window, as in PowerPoint's Normal view, so it is kept;
 * the scroll position was chosen for the previous slide's content and says
 * nothing about the next one's. Top-left rather than centred because that is
 * where a slide's content starts in reading order (its title, in most
 * layouts), and because it needs no layout: a centred position depends on the
 * scroll range at the new slide's size, which may not be laid out yet when the
 * binding reacts. When nothing overflows the viewport there is nothing to
 * scroll and this changes nothing.
 */
export function resetViewportScroll(viewport: Pick<HTMLElement, 'scrollLeft' | 'scrollTop'>): void {
	viewport.scrollLeft = 0;
	viewport.scrollTop = 0;
}

/**
 * Elements whose own Space handling (typing, pressing) must win over pan. Not
 * `isEditorControlTarget` from `editor-key-target`: that treats everything
 * inside the slide viewport as the canvas, and a focused button there (a chart
 * or table quick action) must still be pressed by Space.
 */
const SPACE_OWNER_SELECTOR = [
	'input',
	'textarea',
	'select',
	'button',
	'a[href]',
	// Space toggles a disclosure and plays or pauses media with native controls.
	'summary',
	'video[controls]',
	'audio[controls]',
	'[role="button"]',
	'[role="checkbox"]',
	'[role="combobox"]',
	'[role="menuitem"]',
	'[role="menuitemcheckbox"]',
	'[role="menuitemradio"]',
	'[role="option"]',
	'[role="radio"]',
	'[role="searchbox"]',
	'[role="switch"]',
	'[role="tab"]',
	'[role="textbox"]',
].join(',');

/**
 * Whether a key press comes from an element that uses Space itself (text entry,
 * buttons, links). Reads the event's real target, `composedPath()[0]`: at a
 * document listener a control inside an open shadow root (the status bar's zoom
 * buttons) is otherwise retargeted to its host element and looks like nothing.
 */
export function eventOwnsSpaceKey(event: Event): boolean {
	return targetOwnsSpaceKey(event.composedPath?.()[0] ?? event.target);
}

/** Whether the focused element uses Space itself (text entry, buttons, links). */
export function targetOwnsSpaceKey(target: EventTarget | null): boolean {
	const element = target as (Element & { isContentEditable?: boolean }) | null;
	if (!element || typeof element.closest !== 'function') {
		return false;
	}
	return element.isContentEditable === true || element.closest(SPACE_OWNER_SELECTOR) !== null;
}

/** Events that report how many fingers are on the viewport. */
const TOUCH_EVENTS = ['touchstart', 'touchend', 'touchcancel'] as const;

/** Safari's pinch gesture event; not in the DOM typings. */
interface PinchGestureEvent extends UIEvent {
	readonly scale: number;
	readonly clientX: number;
	readonly clientY: number;
}

export interface ViewportNavigationOptions {
	/** The user's zoom factor (1 = fit), read live so a burst of wheel events compounds. */
	getZoom(): number;
	/** Apply a user zoom factor; the binding's zoom store clamps it to its range. */
	setZoom(zoom: number): void;
	/**
	 * True when a plain primary-button drag has nothing else to do and should
	 * pan, such as a read-only view. Defaults to false: in an editor a plain
	 * drag selects.
	 */
	dragPans?(): boolean;
	/** False while the viewport must leave the wheel and pointer alone. Defaults to true. */
	isActive?(): boolean;
	/** Frame scheduler; defaults to the viewport window's `requestAnimationFrame`. */
	requestFrame?(callback: () => void): number;
	/** Cancels a frame from {@link ViewportNavigationOptions.requestFrame}. */
	cancelFrame?(handle: number): void;
	/** Millisecond clock for telling gestures apart; defaults to `performance.now()`. */
	now?(): number;
}

interface PendingAnchor {
	anchor: ZoomAnchor;
	/** Slide width when the anchor was taken: the layout to wait past. */
	width: number;
	frames: number;
}

interface PanState {
	pointerId: number;
	startX: number;
	startY: number;
	left: number;
	top: number;
	/** False until a threshold drag has moved far enough to count as a pan. */
	started: boolean;
	/** A threshold drag (read-only plain drag) rather than a middle or Space drag. */
	thresholded: boolean;
}

/**
 * Zoom at the pointer and pan by dragging in one slide viewport. Returns the
 * detach function.
 */
export function attachViewportNavigation(
	viewport: HTMLElement,
	options: ViewportNavigationOptions,
): () => void {
	const doc = viewport.ownerDocument;
	const win = doc.defaultView;
	const requestFrame =
		options.requestFrame ?? ((callback: () => void) => win?.requestAnimationFrame(callback) ?? 0);
	const cancelFrame =
		options.cancelFrame ?? ((handle: number) => win?.cancelAnimationFrame(handle));
	const now = options.now ?? (() => win?.performance.now() ?? Date.now());
	const isActive = (): boolean => options.isActive?.() ?? true;
	const wheelBuffer = createWheelStepBuffer();
	const originalCursor = viewport.style.cursor;

	let pending: PendingAnchor | null = null;
	/** The current gesture's slide point and when it last zoomed. */
	let gesture: { anchor: ZoomAnchor; time: number } | null = null;
	let frame = 0;
	let gestureStartZoom: number | null = null;
	/** Fingers on the viewport, from the touch events Safari sends alongside gestures. */
	let fingers = 0;
	let hovering = false;
	let spaceHeld = false;
	let pan: PanState | null = null;

	// Looked up on each use, so a re-rendered stage needs no re-attach.
	const findStage = (): Element | null => viewport.querySelector(VIEWPORT_STAGE_SELECTOR);

	const applyPending = (): void => {
		frame = 0;
		if (!pending) {
			return;
		}
		const rect = findStage()?.getBoundingClientRect();
		if (!rect || rect.width === pending.width) {
			pending.frames += 1;
			if (pending.frames < MAX_LAYOUT_FRAMES) {
				frame = requestFrame(applyPending);
			} else {
				pending = null;
			}
			return;
		}
		const next = anchoredScrollPosition(
			{ left: viewport.scrollLeft, top: viewport.scrollTop },
			rect,
			pending.anchor,
		);
		pending = null;
		viewport.scrollLeft = next.left;
		viewport.scrollTop = next.top;
	};

	const gestureAnchor = (
		rect: ViewportRect | undefined,
		clientX: number,
		clientY: number,
	): ZoomAnchor | null => {
		const time = now();
		const continues =
			gesture !== null &&
			time - gesture.time <= ANCHOR_REUSE_MS &&
			Math.abs(clientX - gesture.anchor.clientX) <= ANCHOR_REUSE_DISTANCE_PX &&
			Math.abs(clientY - gesture.anchor.clientY) <= ANCHOR_REUSE_DISTANCE_PX;
		const anchor = continues
			? { ...gesture!.anchor, clientX, clientY }
			: rect
				? zoomAnchorAt(rect, clientX, clientY)
				: null;
		gesture = anchor ? { anchor, time } : null;
		return anchor;
	};

	const zoomAt = (zoom: number, clientX: number, clientY: number): void => {
		const before = options.getZoom();
		const rect = findStage()?.getBoundingClientRect();
		const anchor = gestureAnchor(rect, clientX, clientY);
		if (rect && pending && pending.width === rect.width) {
			// Still waiting for the last step's layout: the slide has not moved yet.
			if (anchor) {
				pending.anchor = anchor;
			}
		} else {
			pending = rect && anchor ? { anchor, width: rect.width, frames: 0 } : null;
		}
		options.setZoom(zoom);
		if (options.getZoom() === before) {
			// Clamped at the end of the range: nothing will move.
			pending = null;
			return;
		}
		if (pending && frame === 0) {
			frame = requestFrame(applyPending);
		}
	};

	const onWheel = (event: WheelEvent): void => {
		if (!isActive()) {
			return;
		}
		const mapped = mapEditingWheel(event, wheelBuffer, true);
		if (mapped.intent !== 'zoom') {
			return;
		}
		event.preventDefault();
		zoomAt(options.getZoom() + mapped.deltaScale, event.clientX, event.clientY);
	};

	const onTouches = (event: TouchEvent): void => {
		fingers = event.touches.length;
	};
	const onGestureStart = (event: Event): void => {
		// iOS Safari also fires these for a two-finger touch pinch, which the
		// binding's touch recognisers already own; only a trackpad pinch (no
		// fingers on the screen) is zoomed here.
		if (!isActive() || fingers > 0) {
			return;
		}
		event.preventDefault();
		gestureStartZoom = options.getZoom();
	};
	const onGestureChange = (event: Event): void => {
		if (gestureStartZoom === null) {
			return;
		}
		event.preventDefault();
		const gesture = event as PinchGestureEvent;
		zoomAt(gestureStartZoom * gesture.scale, gesture.clientX, gesture.clientY);
	};
	const onGestureEnd = (event: Event): void => {
		if (gestureStartZoom === null) {
			return;
		}
		event.preventDefault();
		gestureStartZoom = null;
	};

	const updateCursor = (): void => {
		if (pan?.started) {
			viewport.style.cursor = 'grabbing';
		} else if (spaceHeld && hovering) {
			viewport.style.cursor = 'grab';
		} else {
			viewport.style.cursor = originalCursor;
		}
	};

	const onPointerEnter = (): void => {
		hovering = true;
		updateCursor();
	};
	const onPointerLeave = (): void => {
		hovering = false;
		updateCursor();
	};

	const onKeyDown = (event: KeyboardEvent): void => {
		if (event.key !== ' ' || event.ctrlKey || event.metaKey || event.altKey) {
			return;
		}
		if (spaceHeld) {
			// Auto-repeat while held must not scroll the page or press anything.
			event.preventDefault();
			return;
		}
		if (!hovering || !isActive() || eventOwnsSpaceKey(event) || !viewportCanPan(viewport)) {
			return;
		}
		event.preventDefault();
		spaceHeld = true;
		updateCursor();
	};
	const releaseSpace = (): void => {
		if (!spaceHeld) {
			return;
		}
		spaceHeld = false;
		updateCursor();
	};
	const onKeyUp = (event: KeyboardEvent): void => {
		if (event.key === ' ' && spaceHeld) {
			event.preventDefault();
			releaseSpace();
		}
	};

	const suppressNextClick = (): void => {
		const onClick = (event: MouseEvent): void => {
			event.preventDefault();
			event.stopPropagation();
		};
		viewport.addEventListener('click', onClick, { capture: true, once: true });
		// A drag that ends outside the viewport produces no click on it.
		win?.setTimeout(() => viewport.removeEventListener('click', onClick, true), 0);
	};

	const startPanning = (event: PointerEvent): void => {
		if (!pan) {
			return;
		}
		pan.started = true;
		try {
			viewport.setPointerCapture(event.pointerId);
		} catch {
			// The pointer may already be gone; panning still follows its moves.
		}
		updateCursor();
	};

	const onPointerDown = (event: PointerEvent): void => {
		if (pan || !isActive() || !viewportCanPan(viewport)) {
			return;
		}
		const middle = event.button === 1;
		const spaceDrag = event.button === 0 && spaceHeld;
		const plainDrag =
			event.button === 0 &&
			!spaceHeld &&
			event.pointerType !== 'touch' &&
			(options.dragPans?.() ?? false);
		if (!middle && !spaceDrag && !plainDrag) {
			return;
		}
		pan = {
			pointerId: event.pointerId,
			startX: event.clientX,
			startY: event.clientY,
			left: viewport.scrollLeft,
			top: viewport.scrollTop,
			started: false,
			thresholded: plainDrag,
		};
		if (!plainDrag) {
			// Middle and Space drags are pans from the first pixel: keep the editor
			// from starting a marquee or a shape drag, and the browser from
			// starting its own middle-button autoscroll.
			event.preventDefault();
			event.stopPropagation();
			startPanning(event);
		}
	};

	const onPointerMove = (event: PointerEvent): void => {
		if (!pan || event.pointerId !== pan.pointerId) {
			return;
		}
		const dx = event.clientX - pan.startX;
		const dy = event.clientY - pan.startY;
		if (!pan.started) {
			if (Math.hypot(dx, dy) < DRAG_PAN_THRESHOLD_PX) {
				return;
			}
			startPanning(event);
		}
		event.preventDefault();
		viewport.scrollLeft = pan.left - dx;
		viewport.scrollTop = pan.top - dy;
	};

	const endPan = (event: PointerEvent): void => {
		if (!pan || event.pointerId !== pan.pointerId) {
			return;
		}
		const { started, thresholded } = pan;
		pan = null;
		if (started) {
			try {
				viewport.releasePointerCapture(event.pointerId);
			} catch {
				// Already released.
			}
			if (thresholded && event.type === 'pointerup') {
				suppressNextClick();
			}
		}
		updateCursor();
	};

	const onWindowBlur = (): void => {
		releaseSpace();
	};

	viewport.addEventListener('wheel', onWheel, { passive: false });
	for (const type of TOUCH_EVENTS) {
		viewport.addEventListener(type, onTouches, { passive: true });
	}
	viewport.addEventListener('gesturestart', onGestureStart);
	viewport.addEventListener('gesturechange', onGestureChange);
	viewport.addEventListener('gestureend', onGestureEnd);
	viewport.addEventListener('pointerenter', onPointerEnter);
	viewport.addEventListener('pointerleave', onPointerLeave);
	viewport.addEventListener('pointerdown', onPointerDown, true);
	viewport.addEventListener('pointermove', onPointerMove, true);
	viewport.addEventListener('pointerup', endPan, true);
	viewport.addEventListener('pointercancel', endPan, true);
	doc.addEventListener('keydown', onKeyDown, true);
	doc.addEventListener('keyup', onKeyUp, true);
	win?.addEventListener('blur', onWindowBlur);

	return () => {
		if (frame !== 0) {
			cancelFrame(frame);
			frame = 0;
		}
		pending = null;
		pan = null;
		spaceHeld = false;
		viewport.style.cursor = originalCursor;
		viewport.removeEventListener('wheel', onWheel);
		for (const type of TOUCH_EVENTS) {
			viewport.removeEventListener(type, onTouches);
		}
		viewport.removeEventListener('gesturestart', onGestureStart);
		viewport.removeEventListener('gesturechange', onGestureChange);
		viewport.removeEventListener('gestureend', onGestureEnd);
		viewport.removeEventListener('pointerenter', onPointerEnter);
		viewport.removeEventListener('pointerleave', onPointerLeave);
		viewport.removeEventListener('pointerdown', onPointerDown, true);
		viewport.removeEventListener('pointermove', onPointerMove, true);
		viewport.removeEventListener('pointerup', endPan, true);
		viewport.removeEventListener('pointercancel', endPan, true);
		doc.removeEventListener('keydown', onKeyDown, true);
		doc.removeEventListener('keyup', onKeyUp, true);
		win?.removeEventListener('blur', onWindowBlur);
	};
}
