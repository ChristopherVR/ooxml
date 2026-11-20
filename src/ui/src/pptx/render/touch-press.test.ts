import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	TOUCH_DOUBLE_TAP_DISTANCE_PX,
	TOUCH_DOUBLE_TAP_MS,
	TOUCH_LONG_PRESS_MOVE_PX,
	TOUCH_LONG_PRESS_MS,
	createTouchPressRecognizer,
	isTouchPressPointer,
} from './touch-press';
import type { TouchPressPointer } from './touch-press';

function pointer(
	clientX: number,
	clientY: number,
	timeStamp: number,
	extra: Partial<TouchPressPointer> = {},
): TouchPressPointer {
	return { pointerId: 1, pointerType: 'touch', clientX, clientY, timeStamp, ...extra };
}

describe('createTouchPressRecognizer', () => {
	beforeEach(() => {
		vi.useFakeTimers();
	});
	afterEach(() => {
		vi.useRealTimers();
	});

	it('reports a long-press with the press coordinates and target', () => {
		const onLongPress = vi.fn();
		const r = createTouchPressRecognizer<string>({ onLongPress });
		r.onPointerDown(pointer(40, 50, 0), 'slide-1');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS - 1);
		expect(onLongPress).not.toHaveBeenCalled();
		vi.advanceTimersByTime(1);
		expect(onLongPress).toHaveBeenCalledWith({ clientX: 40, clientY: 50, target: 'slide-1' });
		expect(r.onPointerUp(pointer(40, 50, TOUCH_LONG_PRESS_MS + 20))).toBe('long-press');
	});

	it('cancels the long-press when the finger moves past the tolerance', () => {
		const onLongPress = vi.fn();
		const r = createTouchPressRecognizer<string>({ onLongPress });
		r.onPointerDown(pointer(0, 0, 0), 'a');
		r.onPointerMove(pointer(TOUCH_LONG_PRESS_MOVE_PX, 0, 50));
		r.onPointerMove(pointer(0, TOUCH_LONG_PRESS_MOVE_PX + 1, 60));
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS * 2);
		expect(onLongPress).not.toHaveBeenCalled();
	});

	it('cancels the long-press on release, cancel and a second finger', () => {
		const onLongPress = vi.fn();
		const r = createTouchPressRecognizer<string>({ onLongPress });
		r.onPointerDown(pointer(0, 0, 0), 'a');
		r.onPointerUp(pointer(0, 0, 100));
		r.onPointerDown(pointer(0, 0, 1000), 'a');
		r.cancel();
		r.onPointerDown(pointer(0, 0, 2000), 'a');
		r.onPointerDown(pointer(30, 30, 2010, { pointerId: 2, isPrimary: false }), 'a');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS * 2);
		expect(onLongPress).not.toHaveBeenCalled();
	});

	it('ignores the mouse, which fires real contextmenu and dblclick events', () => {
		const onLongPress = vi.fn();
		const onDoubleTap = vi.fn();
		const r = createTouchPressRecognizer<string>({ onLongPress, onDoubleTap });
		const mouse = { pointerType: 'mouse' };
		r.onPointerDown(pointer(0, 0, 0, mouse), 'a');
		r.onPointerUp(pointer(0, 0, 10, mouse));
		r.onPointerDown(pointer(0, 0, 50, mouse), 'a');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		expect(onLongPress).not.toHaveBeenCalled();
		expect(r.onPointerUp(pointer(0, 0, 600, mouse))).toBeNull();
		expect(onDoubleTap).not.toHaveBeenCalled();
		expect(isTouchPressPointer('pen')).toBe(true);
		expect(isTouchPressPointer('mouse')).toBe(false);
	});

	it('reports a double-tap for two quick taps close together on one target', () => {
		const onDoubleTap = vi.fn();
		const r = createTouchPressRecognizer<string>({ onDoubleTap });
		r.onPointerDown(pointer(10, 10, 0), 'card');
		expect(r.onPointerUp(pointer(10, 10, 60))).toBeNull();
		r.onPointerDown(pointer(20, 20, TOUCH_DOUBLE_TAP_MS), 'card');
		expect(r.onPointerUp(pointer(20, 20, TOUCH_DOUBLE_TAP_MS + 40))).toBe('double-tap');
		expect(onDoubleTap).toHaveBeenCalledWith({ clientX: 20, clientY: 20, target: 'card' });
		// A third tap starts over rather than making a second double-tap.
		r.onPointerDown(pointer(20, 20, TOUCH_DOUBLE_TAP_MS + 100), 'card');
		expect(r.onPointerUp(pointer(20, 20, TOUCH_DOUBLE_TAP_MS + 140))).toBeNull();
		expect(onDoubleTap).toHaveBeenCalledTimes(1);
	});

	it('needs the second tap in time, near the first and on the same target', () => {
		const onDoubleTap = vi.fn();
		const r = createTouchPressRecognizer<string>({ onDoubleTap });
		const tap = (x: number, t: number, target: string) => {
			r.onPointerDown(pointer(x, 0, t), target);
			return r.onPointerUp(pointer(x, 0, t + 30));
		};
		tap(0, 0, 'a');
		expect(tap(0, TOUCH_DOUBLE_TAP_MS + 1, 'a')).toBeNull();
		expect(tap(TOUCH_DOUBLE_TAP_DISTANCE_PX + 1, TOUCH_DOUBLE_TAP_MS + 100, 'a')).toBeNull();
		expect(tap(TOUCH_DOUBLE_TAP_DISTANCE_PX + 1, TOUCH_DOUBLE_TAP_MS + 200, 'b')).toBeNull();
		expect(onDoubleTap).not.toHaveBeenCalled();
	});

	it('does not count a drag or a long-press as a tap', () => {
		const onDoubleTap = vi.fn();
		const r = createTouchPressRecognizer<string>({ onDoubleTap, onLongPress: () => undefined });
		r.onPointerDown(pointer(0, 0, 0), 'a');
		r.onPointerMove(pointer(40, 0, 20));
		r.onPointerUp(pointer(40, 0, 40));
		r.onPointerDown(pointer(0, 0, 100), 'a');
		r.onPointerUp(pointer(0, 0, 130));
		expect(onDoubleTap).not.toHaveBeenCalled();
		r.onPointerDown(pointer(0, 0, 200), 'a');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		r.onPointerUp(pointer(0, 0, 800));
		r.onPointerDown(pointer(0, 0, 850), 'a');
		r.onPointerUp(pointer(0, 0, 880));
		expect(onDoubleTap).not.toHaveBeenCalled();
	});
});
