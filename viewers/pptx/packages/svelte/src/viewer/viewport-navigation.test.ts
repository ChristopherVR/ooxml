// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { viewportNavigation } from './viewport-navigation';
import type { ViewportNavigationParams } from './viewport-navigation';

function ctrlWheel(): WheelEvent {
	const event = new WheelEvent('wheel', {
		deltaY: -100,
		cancelable: true,
		clientX: 10,
		clientY: 10,
	});
	// Happy DOM's WheelEvent omits the inherited mouse modifier properties.
	Object.defineProperty(event, 'ctrlKey', { value: true });
	return event;
}

function setup(overrides: Partial<ViewportNavigationParams> = {}) {
	const node = document.createElement('div');
	node.innerHTML = '<div aria-roledescription="slide"></div>';
	document.body.appendChild(node);
	let scroll = 0;
	Object.defineProperties(node, {
		clientWidth: { configurable: true, get: () => 100 },
		clientHeight: { configurable: true, get: () => 100 },
		scrollWidth: { configurable: true, get: () => 300 },
		scrollHeight: { configurable: true, get: () => 300 },
		scrollLeft: {
			configurable: true,
			get: () => scroll,
			set: (value: number) => {
				scroll = Math.max(0, value);
			},
		},
		scrollTop: { configurable: true, get: () => 0, set: () => {} },
	});
	vi.spyOn(node, 'setPointerCapture').mockImplementation(() => {});
	vi.spyOn(node, 'releasePointerCapture').mockImplementation(() => {});
	let zoom = 1;
	const params: ViewportNavigationParams = {
		getScale: () => zoom,
		setScale: (next) => (zoom = next),
		dragPans: true,
		active: true,
		...overrides,
	};
	const action = viewportNavigation(node, params);
	const drag = (): number => {
		scroll = 100;
		const pointer = (type: string, clientX: number) => {
			const event = new MouseEvent(type, {
				bubbles: true,
				cancelable: true,
				button: 0,
				clientX,
				clientY: 10,
			});
			Object.defineProperties(event, { pointerId: { value: 1 }, pointerType: { value: 'mouse' } });
			node.dispatchEvent(event);
		};
		pointer('pointerdown', 200);
		pointer('pointermove', 150);
		pointer('pointerup', 150);
		return scroll;
	};
	return { node, params, action, drag, zoom: () => zoom };
}

describe('viewportNavigation action', () => {
	let env: ReturnType<typeof setup> | null = null;
	afterEach(() => {
		env?.action.destroy();
		env?.node.remove();
		env = null;
	});

	it('zooms through the binding scale with Ctrl+wheel', () => {
		env = setup();
		const event = ctrlWheel();
		env.node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeTruthy();
		expect(env.zoom()).toBeGreaterThan(1);
	});

	it('pans with a plain drag only while dragPans is set, following updates', () => {
		env = setup();
		expect(env.drag()).toBe(150);
		env.action.update({ ...env.params, dragPans: false });
		expect(env.drag()).toBe(100);
	});

	it('leaves the wheel alone while inactive', () => {
		env = setup({ active: false });
		const event = ctrlWheel();
		env.node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeFalsy();
		expect(env.zoom()).toBe(1);
	});

	it('detaches on destroy', () => {
		env = setup();
		env.action.destroy();
		const event = ctrlWheel();
		env.node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeFalsy();
		expect(env.zoom()).toBe(1);
	});
});
