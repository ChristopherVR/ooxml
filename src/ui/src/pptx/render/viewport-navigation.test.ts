// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
	DRAG_PAN_THRESHOLD_PX,
	anchoredScrollPosition,
	attachViewportNavigation,
	resetViewportScroll,
	targetOwnsSpaceKey,
	viewportCanPan,
	zoomAnchorAt,
} from './viewport-navigation';
import type { ViewportNavigationOptions, ViewportRect } from './viewport-navigation';

describe('zoom anchoring maths', () => {
	it('keeps the slide point under the pointer across a zoom', () => {
		const before: ViewportRect = { left: 100, top: 50, width: 400, height: 225 };
		const anchor = zoomAnchorAt(before, 200, 100);
		expect(anchor).toEqual({ fx: 0.25, fy: 50 / 225, clientX: 200, clientY: 100 });

		// Twice as large, laid out from the same origin with no scroll yet.
		const after: ViewportRect = { left: 100, top: 50, width: 800, height: 450 };
		const scroll = anchoredScrollPosition({ left: 0, top: 0 }, after, anchor!);
		expect(scroll).toEqual({ left: 100, top: 50 });
		// Scrolling by that much moves the slide so the point is back under the pointer.
		expect(after.left - scroll.left + anchor!.fx * after.width).toBe(200);
		expect(after.top - scroll.top + anchor!.fy * after.height).toBeCloseTo(100);
	});

	it('has no anchor before the slide is laid out', () => {
		expect(zoomAnchorAt({ left: 0, top: 0, width: 0, height: 0 }, 10, 10)).toBeNull();
	});

	it('pans only when the content is larger than the viewport', () => {
		expect(
			viewportCanPan({ scrollWidth: 800, scrollHeight: 450, clientWidth: 800, clientHeight: 450 }),
		).toBe(false);
		expect(
			viewportCanPan({ scrollWidth: 1600, scrollHeight: 450, clientWidth: 800, clientHeight: 450 }),
		).toBe(true);
		expect(
			viewportCanPan({ scrollWidth: 800, scrollHeight: 900, clientWidth: 800, clientHeight: 450 }),
		).toBe(true);
	});

	it('shows a new slide from its top-left corner', () => {
		const viewport = { scrollLeft: 320, scrollTop: 180 };
		resetViewportScroll(viewport);
		expect(viewport).toStrictEqual({ scrollLeft: 0, scrollTop: 0 });
	});

	it('leaves Space to text entry, buttons, links, disclosures, media controls and toggles', () => {
		document.body.innerHTML = [
			'<input id="i"><button id="b"></button><a id="a" href="#"></a><div id="d" tabindex="0"></div>',
			'<details><summary id="s">More</summary></details>',
			'<video id="v" controls></video><audio id="au" controls></audio><video id="bare"></video>',
			'<div id="r" role="radio"></div><div id="mc" role="menuitemcheckbox"></div>',
			'<div id="mr" role="menuitemradio"></div><div id="sb" role="searchbox"></div>',
			'<div id="cb" role="combobox"></div>',
		].join('');
		for (const id of ['i', 'b', 'a', 's', 'v', 'au', 'r', 'mc', 'mr', 'sb', 'cb']) {
			expect(targetOwnsSpaceKey(document.getElementById(id)), id).toBe(true);
		}
		// A video without controls has no Space action of its own.
		expect(targetOwnsSpaceKey(document.getElementById('bare'))).toBe(false);
		const editable = document.createElement('div');
		Object.defineProperty(editable, 'isContentEditable', { value: true });
		expect(targetOwnsSpaceKey(editable)).toBe(true);
		expect(targetOwnsSpaceKey(document.getElementById('d'))).toBe(false);
		expect(targetOwnsSpaceKey(null)).toBe(false);
	});
});

/** A viewport with a fake layout: the stage rect follows the zoom like a real one. */
function setup(overrides: Partial<ViewportNavigationOptions> = {}) {
	document.body.innerHTML =
		'<div data-pptx-viewport><div><div aria-roledescription="slide"></div></div></div>';
	const viewport = document.querySelector<HTMLElement>('[data-pptx-viewport]')!;
	const stage = viewport.querySelector<HTMLElement>('[aria-roledescription="slide"]')!;
	const layout = { zoom: 1, laidOutZoom: 1, scrollLeft: 0, scrollTop: 0 };
	const base = { left: 100, top: 50, width: 400, height: 225 };
	// The stage is laid out at `laidOutZoom`, which only catches up with `zoom` on `layoutFrame()`,
	// like a binding that re-renders after the store changes.
	stage.getBoundingClientRect = () =>
		({
			left: base.left - layout.scrollLeft,
			top: base.top - layout.scrollTop,
			width: base.width * layout.laidOutZoom,
			height: base.height * layout.laidOutZoom,
		}) as DOMRect;
	Object.defineProperties(viewport, {
		clientWidth: { configurable: true, get: () => 600 },
		clientHeight: { configurable: true, get: () => 400 },
		scrollWidth: {
			configurable: true,
			get: () => Math.max(600, base.width * layout.laidOutZoom + 200),
		},
		scrollHeight: {
			configurable: true,
			get: () => Math.max(400, base.height * layout.laidOutZoom + 100),
		},
		scrollLeft: {
			configurable: true,
			get: () => layout.scrollLeft,
			set: (value: number) => {
				layout.scrollLeft = Math.max(0, value);
			},
		},
		scrollTop: {
			configurable: true,
			get: () => layout.scrollTop,
			set: (value: number) => {
				layout.scrollTop = Math.max(0, value);
			},
		},
	});
	viewport.setPointerCapture = vi.fn();
	viewport.releasePointerCapture = vi.fn();
	const frames: (() => void)[] = [];
	const setZoom = vi.fn((zoom: number) => {
		layout.zoom = Math.min(Math.max(zoom, 0.2), 5);
	});
	const detach = attachViewportNavigation(viewport, {
		getZoom: () => layout.zoom,
		setZoom,
		requestFrame: (callback) => frames.push(callback),
		cancelFrame: vi.fn(),
		...overrides,
	});
	const layoutFrame = () => {
		layout.laidOutZoom = layout.zoom;
		frames.splice(0).forEach((callback) => callback());
	};
	return { viewport, stage, layout, setZoom, detach, frames, layoutFrame };
}

function pointer(
	type: string,
	init: { button?: number; clientX?: number; clientY?: number; pointerType?: string },
) {
	const event = new MouseEvent(type, { bubbles: true, cancelable: true, ...init });
	Object.defineProperties(event, {
		pointerId: { value: 1 },
		pointerType: { value: init.pointerType ?? 'mouse' },
	});
	return event;
}

describe('attachViewportNavigation: zoom at the pointer', () => {
	let env: ReturnType<typeof setup>;
	beforeEach(() => {
		env = setup();
	});
	afterEach(() => env.detach());

	it('zooms on Ctrl+wheel and scrolls so the slide point stays under the pointer', () => {
		const wheel = new WheelEvent('wheel', {
			ctrlKey: true,
			deltaY: -100,
			clientX: 200,
			clientY: 100,
			cancelable: true,
		});
		env.viewport.dispatchEvent(wheel);

		expect(wheel.defaultPrevented).toBe(true);
		expect(env.setZoom).toHaveBeenCalledTimes(1);
		const zoom = env.layout.zoom;
		expect(zoom).toBeGreaterThan(1);

		env.layoutFrame();
		const rect = env.stage.getBoundingClientRect();
		// The point a quarter of the way across the slide is still at x=200.
		expect(rect.left + 0.25 * rect.width).toBeCloseTo(200);
		expect(rect.top + (50 / 225) * rect.height).toBeCloseTo(100);
	});

	it('zooms on Cmd+wheel too', () => {
		env.viewport.dispatchEvent(
			new WheelEvent('wheel', { metaKey: true, deltaY: -100, cancelable: true }),
		);
		expect(env.setZoom).toHaveBeenCalledTimes(1);
	});

	it('keeps one slide point through a burst of wheel events before the layout catches up', () => {
		for (let i = 0; i < 3; i += 1) {
			env.viewport.dispatchEvent(
				new WheelEvent('wheel', {
					ctrlKey: true,
					deltaY: -100,
					clientX: 200,
					clientY: 100,
					cancelable: true,
				}),
			);
		}
		expect(env.frames).toHaveLength(1);
		env.layoutFrame();
		const rect = env.stage.getBoundingClientRect();
		expect(rect.left + 0.25 * rect.width).toBeCloseTo(200);
	});

	it("keeps the gesture's first slide point while the zoom continues, even if the slide drifted", () => {
		env.detach();
		const clock = { t: 0 };
		env = setup({ now: () => clock.t });
		const wheelAt = () =>
			env.viewport.dispatchEvent(
				new WheelEvent('wheel', {
					ctrlKey: true,
					deltaY: -100,
					clientX: 200,
					clientY: 100,
					cancelable: true,
				}),
			);
		wheelAt();
		env.layoutFrame();
		// A centred slide cannot scroll to hold the point: model that as the scroll snapping back.
		env.layout.scrollLeft = 0;
		clock.t = 100;
		wheelAt();
		env.layoutFrame();
		const rect = env.stage.getBoundingClientRect();
		expect(rect.left + 0.25 * rect.width).toBeCloseTo(200);
	});

	it('starts a new gesture with the point under the pointer after a pause', () => {
		env.detach();
		const clock = { t: 0 };
		env = setup({ now: () => clock.t });
		const wheelAt = () =>
			env.viewport.dispatchEvent(
				new WheelEvent('wheel', {
					ctrlKey: true,
					deltaY: -100,
					clientX: 200,
					clientY: 100,
					cancelable: true,
				}),
			);
		wheelAt();
		env.layoutFrame();
		env.layout.scrollLeft = 0;
		const drifted = env.stage.getBoundingClientRect();
		const fx = (200 - drifted.left) / drifted.width;
		clock.t = 1000;
		wheelAt();
		env.layoutFrame();
		const rect = env.stage.getBoundingClientRect();
		expect(rect.left + fx * rect.width).toBeCloseTo(200);
		expect(fx).not.toBeCloseTo(0.25);
	});

	it('waits for the slide to be laid out at the new zoom before scrolling', () => {
		env.viewport.dispatchEvent(
			new WheelEvent('wheel', {
				ctrlKey: true,
				deltaY: -100,
				clientX: 200,
				clientY: 100,
				cancelable: true,
			}),
		);
		// A frame where the binding has not re-rendered yet: nothing moves, it tries again.
		env.frames.splice(0).forEach((callback) => callback());
		expect(env.layout.scrollLeft).toBe(0);
		expect(env.frames).toHaveLength(1);
		env.layoutFrame();
		expect(env.layout.scrollLeft).toBeGreaterThan(0);
	});

	it('leaves a plain wheel to scroll natively', () => {
		const wheel = new WheelEvent('wheel', { deltaY: 100, cancelable: true });
		env.viewport.dispatchEvent(wheel);
		expect(wheel.defaultPrevented).toBe(false);
		expect(env.setZoom).not.toHaveBeenCalled();
	});

	it('does not scroll when the zoom is already at its limit', () => {
		env.layout.zoom = 5;
		env.viewport.dispatchEvent(
			new WheelEvent('wheel', {
				ctrlKey: true,
				deltaY: -100,
				clientX: 200,
				clientY: 100,
				cancelable: true,
			}),
		);
		expect(env.frames).toHaveLength(0);
		expect(env.layout.scrollLeft).toBe(0);
	});

	it('zooms on a Safari pinch gesture around the gesture point', () => {
		const start = new UIEvent('gesturestart', { cancelable: true });
		env.viewport.dispatchEvent(start);
		const change = new UIEvent('gesturechange', { cancelable: true });
		Object.defineProperties(change, {
			scale: { value: 2 },
			clientX: { value: 200 },
			clientY: { value: 100 },
		});
		env.viewport.dispatchEvent(change);

		expect(start.defaultPrevented).toBe(true);
		expect(change.defaultPrevented).toBe(true);
		expect(env.layout.zoom).toBe(2);
		env.layoutFrame();
		const rect = env.stage.getBoundingClientRect();
		expect(rect.left + 0.25 * rect.width).toBeCloseTo(200);
	});

	it('leaves a touch pinch to the touch recognisers (iOS sends gesture events for it too)', () => {
		const touches = (type: string, count: number) => {
			const event = new Event(type);
			Object.defineProperty(event, 'touches', { value: { length: count } });
			return event;
		};
		env.viewport.dispatchEvent(touches('touchstart', 2));
		const start = new UIEvent('gesturestart', { cancelable: true });
		env.viewport.dispatchEvent(start);
		const change = new UIEvent('gesturechange', { cancelable: true });
		Object.defineProperty(change, 'scale', { value: 2 });
		env.viewport.dispatchEvent(change);
		expect(start.defaultPrevented).toBe(false);
		expect(env.setZoom).not.toHaveBeenCalled();

		// Fingers lifted: a trackpad pinch zooms again.
		env.viewport.dispatchEvent(touches('touchend', 0));
		env.viewport.dispatchEvent(new UIEvent('gesturestart', { cancelable: true }));
		env.viewport.dispatchEvent(change);
		expect(env.layout.zoom).toBe(2);
	});

	it('ignores the wheel while inactive', () => {
		env.detach();
		env = setup({ isActive: () => false });
		const wheel = new WheelEvent('wheel', { ctrlKey: true, deltaY: -100, cancelable: true });
		env.viewport.dispatchEvent(wheel);
		expect(wheel.defaultPrevented).toBe(false);
		expect(env.setZoom).not.toHaveBeenCalled();
	});

	it('stops listening when detached', () => {
		env.detach();
		env.viewport.dispatchEvent(
			new WheelEvent('wheel', { ctrlKey: true, deltaY: -100, cancelable: true }),
		);
		expect(env.setZoom).not.toHaveBeenCalled();
	});
});

describe('attachViewportNavigation: drag to pan', () => {
	let env: ReturnType<typeof setup>;
	afterEach(() => env.detach());

	const zoomIn = () => {
		env.layout.zoom = 2;
		env.layout.laidOutZoom = 2;
	};

	it('pans with a middle-button drag and keeps the editor from seeing it', () => {
		env = setup();
		zoomIn();
		const editor = vi.fn();
		env.viewport.firstElementChild!.addEventListener('pointerdown', editor);

		const down = pointer('pointerdown', { button: 1, clientX: 300, clientY: 200 });
		env.viewport.firstElementChild!.dispatchEvent(down);
		env.viewport.dispatchEvent(pointer('pointermove', { button: 1, clientX: 250, clientY: 170 }));

		expect(down.defaultPrevented).toBe(true);
		expect(editor).not.toHaveBeenCalled();
		expect(env.layout).toMatchObject({ scrollLeft: 50, scrollTop: 30 });
		expect(env.viewport.style.cursor).toBe('grabbing');
		env.viewport.dispatchEvent(pointer('pointerup', { button: 1 }));
		expect(env.viewport.style.cursor).toBe('');
	});

	it('pans with Space held while hovering the slide, and not from a text field', () => {
		env = setup();
		zoomIn();
		env.viewport.dispatchEvent(new MouseEvent('pointerenter'));
		const space = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
		document.body.dispatchEvent(space);
		expect(space.defaultPrevented).toBe(true);
		expect(env.viewport.style.cursor).toBe('grab');

		env.viewport.dispatchEvent(pointer('pointerdown', { button: 0, clientX: 300, clientY: 200 }));
		env.viewport.dispatchEvent(pointer('pointermove', { button: 0, clientX: 280, clientY: 200 }));
		expect(env.layout.scrollLeft).toBe(20);
		env.viewport.dispatchEvent(pointer('pointerup', { button: 0 }));
		document.body.dispatchEvent(
			new KeyboardEvent('keyup', { key: ' ', bubbles: true, cancelable: true }),
		);
		expect(env.viewport.style.cursor).toBe('');

		const input = document.createElement('input');
		document.body.append(input);
		const typed = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
		input.dispatchEvent(typed);
		expect(typed.defaultPrevented).toBe(false);
	});

	it('leaves Space to a button inside a shadow root (the status bar)', () => {
		env = setup();
		zoomIn();
		env.viewport.dispatchEvent(new MouseEvent('pointerenter'));
		// A document listener sees a key press from inside a shadow root as coming
		// from the host element; only composedPath() still has the real button.
		const host = document.createElement('pptx-ui-status-bar');
		const button = document.createElement('button');
		document.body.append(host);
		const press = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
		Object.defineProperty(press, 'composedPath', { value: () => [button, host, document.body] });
		host.dispatchEvent(press);

		expect(press.defaultPrevented).toBe(false);
		expect(env.viewport.style.cursor).toBe('');
	});

	it('lets go of Space when the window loses focus', () => {
		// The keyup goes to whatever window has focus then, so it never arrives here.
		env = setup();
		zoomIn();
		env.viewport.dispatchEvent(new MouseEvent('pointerenter'));
		document.body.dispatchEvent(
			new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true }),
		);
		expect(env.viewport.style.cursor).toBe('grab');

		window.dispatchEvent(new Event('blur'));
		expect(env.viewport.style.cursor).toBe('');
		const down = pointer('pointerdown', { button: 0, clientX: 300, clientY: 200 });
		env.viewport.dispatchEvent(down);
		env.viewport.dispatchEvent(pointer('pointermove', { button: 0, clientX: 280, clientY: 200 }));
		expect(down.defaultPrevented).toBe(false);
		expect(env.layout.scrollLeft).toBe(0);
	});

	it('leaves a plain drag to the editor by default', () => {
		env = setup();
		zoomIn();
		const down = pointer('pointerdown', { button: 0, clientX: 300, clientY: 200 });
		env.viewport.dispatchEvent(down);
		env.viewport.dispatchEvent(pointer('pointermove', { button: 0, clientX: 200, clientY: 100 }));
		expect(down.defaultPrevented).toBe(false);
		expect(env.layout.scrollLeft).toBe(0);
	});

	it('pans with a plain drag in a read-only view once past the threshold, then swallows the click', () => {
		env = setup({ dragPans: () => true });
		zoomIn();
		const down = pointer('pointerdown', { button: 0, clientX: 300, clientY: 200 });
		env.viewport.dispatchEvent(down);
		// A press is not a pan, so it must not be cancelled.
		expect(down.defaultPrevented).toBe(false);
		env.viewport.dispatchEvent(
			pointer('pointermove', {
				button: 0,
				clientX: 300 - (DRAG_PAN_THRESHOLD_PX - 1),
				clientY: 200,
			}),
		);
		expect(env.layout.scrollLeft).toBe(0);
		env.viewport.dispatchEvent(pointer('pointermove', { button: 0, clientX: 260, clientY: 190 }));
		expect(env.layout).toMatchObject({ scrollLeft: 40, scrollTop: 10 });
		env.viewport.dispatchEvent(pointer('pointerup', { button: 0 }));

		const clicked = vi.fn();
		env.viewport.firstElementChild!.addEventListener('click', clicked);
		env.viewport.firstElementChild!.dispatchEvent(
			new MouseEvent('click', { bubbles: true, cancelable: true }),
		);
		expect(clicked).not.toHaveBeenCalled();
	});

	it('lets a click through in a read-only view when the pointer barely moved', () => {
		env = setup({ dragPans: () => true });
		zoomIn();
		env.viewport.dispatchEvent(pointer('pointerdown', { button: 0, clientX: 300, clientY: 200 }));
		env.viewport.dispatchEvent(pointer('pointermove', { button: 0, clientX: 301, clientY: 201 }));
		env.viewport.dispatchEvent(pointer('pointerup', { button: 0 }));
		const clicked = vi.fn();
		env.viewport.firstElementChild!.addEventListener('click', clicked);
		env.viewport.firstElementChild!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
		expect(clicked).toHaveBeenCalledTimes(1);
	});

	it('leaves touch to native scrolling', () => {
		env = setup({ dragPans: () => true });
		zoomIn();
		env.viewport.dispatchEvent(
			pointer('pointerdown', { button: 0, clientX: 300, clientY: 200, pointerType: 'touch' }),
		);
		env.viewport.dispatchEvent(
			pointer('pointermove', { button: 0, clientX: 200, clientY: 200, pointerType: 'touch' }),
		);
		expect(env.layout.scrollLeft).toBe(0);
	});

	it('does not pan or claim Space while the slide fits', () => {
		env = setup({ dragPans: () => true });
		env.viewport.dispatchEvent(new MouseEvent('pointerenter'));
		const space = new KeyboardEvent('keydown', { key: ' ', bubbles: true, cancelable: true });
		document.body.dispatchEvent(space);
		expect(space.defaultPrevented).toBe(false);
		const down = pointer('pointerdown', { button: 1, clientX: 300, clientY: 200 });
		env.viewport.dispatchEvent(down);
		expect(down.defaultPrevented).toBe(false);
	});
});
