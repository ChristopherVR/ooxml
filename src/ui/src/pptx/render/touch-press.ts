/**
 * Framework-agnostic long-press and double-tap recogniser for touch and pen.
 *
 * iOS Safari fires no `contextmenu` on a long-press and `dblclick` is unreliable
 * for touch, so every control that relies on those events has no touch path of
 * its own (issue #290). This state machine reads pointer-event-shaped objects and
 * reports:
 *   - a long-press: the pointer held for {@link TOUCH_LONG_PRESS_MS} without
 *     moving more than {@link TOUCH_LONG_PRESS_MOVE_PX};
 *   - a double-tap: two short taps within {@link TOUCH_DOUBLE_TAP_MS} and
 *     {@link TOUCH_DOUBLE_TAP_DISTANCE_PX} of each other.
 *
 * Mouse pointers are ignored: a mouse already fires real `contextmenu` and
 * `dblclick` events. The DOM wiring (listeners, synthetic events, click
 * suppression) is `installTouchPress` in `touch-press-dom.ts`.
 *
 * @module touch-press
 */

/** The hold duration (ms) of a long-press. */
export const TOUCH_LONG_PRESS_MS = 500;

/** Moving further than this (px) during the hold cancels the long-press. */
export const TOUCH_LONG_PRESS_MOVE_PX = 10;

/** The longest gap (ms) between two taps that still makes a double-tap. */
export const TOUCH_DOUBLE_TAP_MS = 300;

/** The furthest (px) the second tap may land from the first. */
export const TOUCH_DOUBLE_TAP_DISTANCE_PX = 24;

/** The subset of a DOM `PointerEvent` the recogniser reads. */
export interface TouchPressPointer {
	readonly pointerId: number;
	readonly pointerType: string;
	readonly clientX: number;
	readonly clientY: number;
	/** Milliseconds; any monotonic clock (the DOM's `timeStamp`). */
	readonly timeStamp: number;
	/** `false` for a secondary pointer of a multi-touch gesture. */
	readonly isPrimary?: boolean;
}

/** A recognised gesture, with where it happened and what it started on. */
export interface TouchPressGesture<T> {
	readonly clientX: number;
	readonly clientY: number;
	/** The value passed to `onPointerDown` with the press (usually its target). */
	readonly target: T;
}

export interface TouchPressCallbacks<T> {
	onLongPress?: (gesture: TouchPressGesture<T>) => void;
	onDoubleTap?: (gesture: TouchPressGesture<T>) => void;
}

export interface TouchPressRecognizer<T> {
	/** Starts tracking a press. `target` is handed back with the gesture. */
	readonly onPointerDown: (event: TouchPressPointer, target: T) => void;
	readonly onPointerMove: (event: TouchPressPointer) => void;
	/** Returns the gesture the release completed, if any. */
	readonly onPointerUp: (event: TouchPressPointer) => 'double-tap' | 'long-press' | null;
	/** A cancelled pointer, a scroll or a native menu: forget the press. */
	readonly cancel: () => void;
}

/** True for the pointer types the recogniser handles (touch and pen). */
export function isTouchPressPointer(pointerType: string): boolean {
	return pointerType === 'touch' || pointerType === 'pen';
}

interface Press<T> {
	readonly pointerId: number;
	readonly x: number;
	readonly y: number;
	readonly time: number;
	readonly target: T;
	moved: boolean;
	longPressed: boolean;
}

interface Tap<T> {
	readonly x: number;
	readonly y: number;
	readonly time: number;
	readonly target: T;
}

/**
 * Creates a recogniser. `sameTarget` decides whether two taps hit the same
 * control (default: identity), so a double-tap never spans two slide cards.
 */
export function createTouchPressRecognizer<T>(
	callbacks: TouchPressCallbacks<T>,
	sameTarget: (a: T, b: T) => boolean = Object.is,
): TouchPressRecognizer<T> {
	let press: Press<T> | null = null;
	let lastTap: Tap<T> | null = null;
	let timer: ReturnType<typeof setTimeout> | null = null;

	const clearTimer = (): void => {
		if (timer !== null) {
			clearTimeout(timer);
			timer = null;
		}
	};

	const cancel = (): void => {
		clearTimer();
		press = null;
	};

	const onPointerDown = (event: TouchPressPointer, target: T): void => {
		if (!isTouchPressPointer(event.pointerType)) {
			return;
		}
		if (event.isPrimary === false) {
			// A second finger: a pinch or a two-finger scroll, never a press.
			cancel();
			lastTap = null;
			return;
		}
		cancel();
		const current: Press<T> = {
			pointerId: event.pointerId,
			x: event.clientX,
			y: event.clientY,
			time: event.timeStamp,
			target,
			moved: false,
			longPressed: false,
		};
		press = current;
		if (callbacks.onLongPress) {
			const onLongPress = callbacks.onLongPress;
			timer = setTimeout(() => {
				timer = null;
				current.longPressed = true;
				lastTap = null;
				onLongPress({ clientX: current.x, clientY: current.y, target: current.target });
			}, TOUCH_LONG_PRESS_MS);
		}
	};

	const onPointerMove = (event: TouchPressPointer): void => {
		if (!press || event.pointerId !== press.pointerId) {
			return;
		}
		const dx = event.clientX - press.x;
		const dy = event.clientY - press.y;
		if (Math.abs(dx) > TOUCH_LONG_PRESS_MOVE_PX || Math.abs(dy) > TOUCH_LONG_PRESS_MOVE_PX) {
			press.moved = true;
			clearTimer();
		}
	};

	const onPointerUp = (event: TouchPressPointer): 'double-tap' | 'long-press' | null => {
		const current = press;
		if (!current || event.pointerId !== current.pointerId) {
			return null;
		}
		cancel();
		if (current.longPressed) {
			return 'long-press';
		}
		if (current.moved || event.timeStamp - current.time > TOUCH_LONG_PRESS_MS) {
			lastTap = null;
			return null;
		}
		const previous = lastTap;
		if (
			previous &&
			current.time - previous.time <= TOUCH_DOUBLE_TAP_MS &&
			Math.hypot(current.x - previous.x, current.y - previous.y) <= TOUCH_DOUBLE_TAP_DISTANCE_PX &&
			sameTarget(previous.target, current.target)
		) {
			lastTap = null;
			if (callbacks.onDoubleTap) {
				callbacks.onDoubleTap({ clientX: current.x, clientY: current.y, target: current.target });
				return 'double-tap';
			}
			return null;
		}
		lastTap = { x: current.x, y: current.y, time: current.time, target: current.target };
		return null;
	};

	return { onPointerDown, onPointerMove, onPointerUp, cancel };
}
