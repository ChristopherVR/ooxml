// @vitest-environment happy-dom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { TOUCH_LONG_PRESS_MS } from './touch-press';
import { TOUCH_PRESS_ATTRIBUTE, installTouchPress, touchPressActions } from './touch-press-dom';

function press(target: Element, type: string, x: number, y: number, pointerType = 'touch') {
	target.dispatchEvent(
		new PointerEvent(type, {
			bubbles: true,
			cancelable: true,
			composed: true,
			pointerId: 7,
			pointerType,
			isPrimary: true,
			clientX: x,
			clientY: y,
		}),
	);
}

function tap(target: Element, x = 5, y = 5) {
	press(target, 'pointerdown', x, y);
	press(target, 'pointerup', x, y);
}

describe('installTouchPress', () => {
	let root: HTMLElement;
	let host: HTMLElement;
	let inner: HTMLElement;
	let uninstall: () => void;

	beforeEach(() => {
		vi.useFakeTimers();
		root = document.createElement('div');
		host = document.createElement('div');
		host.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'menu dblclick');
		inner = document.createElement('span');
		host.appendChild(inner);
		root.appendChild(host);
		document.body.appendChild(root);
		uninstall = installTouchPress(root);
	});
	afterEach(() => {
		uninstall();
		root.remove();
		vi.useRealTimers();
	});

	it('turns a long-press into a contextmenu at the touch point and eats the click', () => {
		const menus: MouseEvent[] = [];
		const clicks = vi.fn();
		host.addEventListener('contextmenu', (e) => menus.push(e));
		host.addEventListener('click', clicks);
		press(inner, 'pointerdown', 30, 40);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		expect(menus).toHaveLength(1);
		expect(menus[0]?.target).toBe(inner);
		expect([menus[0]?.clientX, menus[0]?.clientY]).toEqual([30, 40]);
		press(inner, 'pointerup', 30, 40);
		inner.dispatchEvent(new MouseEvent('click', { bubbles: true }));
		expect(clicks).not.toHaveBeenCalled();
		// The next, unrelated click goes through.
		inner.dispatchEvent(new MouseEvent('click', { bubbles: true }));
		expect(clicks).toHaveBeenCalledTimes(1);
	});

	it('lets every click through before any gesture, even when the clock reads zero', () => {
		// A page clock starts at 0 (and fake timers keep it there), which no deadline may cover.
		const now = vi.spyOn(performance, 'now').mockReturnValue(0);
		try {
			const clicks = vi.fn();
			host.addEventListener('click', clicks);
			inner.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
			expect(clicks).toHaveBeenCalledTimes(1);
		} finally {
			now.mockRestore();
		}
	});

	it('turns a double-tap into a dblclick', () => {
		const dbl = vi.fn();
		host.addEventListener('dblclick', dbl);
		tap(inner);
		tap(inner);
		expect(dbl).toHaveBeenCalledTimes(1);
	});

	it('honours the declared actions and ignores mouse, editable text and other elements', () => {
		const menu = vi.fn();
		const dbl = vi.fn();
		root.addEventListener('contextmenu', menu);
		root.addEventListener('dblclick', dbl);
		host.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'dblclick');
		press(inner, 'pointerdown', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		press(inner, 'pointerup', 0, 0);
		host.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'menu');
		tap(inner);
		tap(inner);
		host.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'menu dblclick');
		press(inner, 'pointerdown', 0, 0, 'mouse');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		inner.setAttribute('contenteditable', 'true');
		press(inner, 'pointerdown', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		inner.removeAttribute('contenteditable');
		const outside = document.createElement('div');
		root.appendChild(outside);
		press(outside, 'pointerdown', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		expect(menu).not.toHaveBeenCalled();
		expect(dbl).not.toHaveBeenCalled();
		expect(touchPressActions(host)).toEqual(new Set(['menu', 'dblclick']));
	});

	it('finds the nearest host for each action, so a nested control keeps the menu', () => {
		const menu = vi.fn();
		const dbl = vi.fn();
		host.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'menu');
		inner.setAttribute(TOUCH_PRESS_ATTRIBUTE, 'dblclick');
		host.addEventListener('contextmenu', menu);
		host.addEventListener('dblclick', dbl);
		press(inner, 'pointerdown', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		press(inner, 'pointerup', 0, 0);
		tap(inner);
		tap(inner);
		expect(menu).toHaveBeenCalledTimes(1);
		expect(dbl).toHaveBeenCalledTimes(1);
	});

	it('cancels on pointercancel', () => {
		const menu = vi.fn();
		host.addEventListener('contextmenu', menu);
		press(inner, 'pointerdown', 0, 0);
		press(inner, 'pointercancel', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		expect(menu).not.toHaveBeenCalled();
	});

	it('stops after teardown', () => {
		uninstall();
		const menu = vi.fn();
		host.addEventListener('contextmenu', menu);
		press(inner, 'pointerdown', 0, 0);
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		expect(menu).not.toHaveBeenCalled();
		uninstall = () => undefined;
	});
});
