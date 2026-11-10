import { mount } from '@vue/test-utils';
import { createViewerZoomStore } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h, nextTick, ref } from 'vue';

import { useViewportNavigation } from './useViewportNavigation';

// Alias keeps the composable out of a `use*`-named call site in plain helpers,
// which the react-hooks linter would otherwise flag as a rules-of-hooks issue.
const runViewportNavigation = useViewportNavigation;

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

async function setup() {
	const store = createViewerZoomStore();
	const viewport = ref<HTMLElement | null>(null);
	const dragPans = ref(true);
	const active = ref(true);
	const slideId = ref<string | undefined>('s1');
	const wrapper = mount(
		defineComponent({
			setup() {
				runViewportNavigation({ viewport, store, dragPans, active, slideId });
				return () => h('main', { ref: viewport }, [h('div', { 'aria-roledescription': 'slide' })]);
			},
		}),
		{ attachTo: document.body },
	);
	// The template ref is set during render; the watcher attaches on the post flush.
	await nextTick();
	const node = viewport.value!;
	let scroll = 0;
	let top = 0;
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
		scrollTop: {
			configurable: true,
			get: () => top,
			set: (value: number) => {
				top = Math.max(0, value);
			},
		},
	});
	vi.spyOn(node, 'setPointerCapture').mockImplementation(() => {});
	vi.spyOn(node, 'releasePointerCapture').mockImplementation(() => {});
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
	return { store, node, dragPans, active, slideId, wrapper, drag };
}

describe('useViewportNavigation', () => {
	let env: Awaited<ReturnType<typeof setup>> | null = null;
	afterEach(() => {
		env?.wrapper.unmount();
		env = null;
	});

	it('zooms the store with Ctrl+wheel in the viewport', async () => {
		env = await setup();
		const event = ctrlWheel();
		env.node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeTruthy();
		expect(env.store.getState().zoom).toBeGreaterThan(1);
	});

	it('pans with a plain drag only while dragPans is set, read live', async () => {
		env = await setup();
		expect(env.drag()).toBe(150);
		env.dragPans.value = false;
		await nextTick();
		expect(env.drag()).toBe(100);
	});

	it('leaves the wheel alone while inactive', async () => {
		env = await setup();
		env.active.value = false;
		const event = ctrlWheel();
		env.node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeFalsy();
		expect(env.store.getState().zoom).toBe(1);
	});

	it('keeps the zoom on another slide and shows it from its top-left corner', async () => {
		env = await setup();
		env.store.dispatch({ type: 'set-zoom', zoom: 0.5 });
		env.node.scrollLeft = 120;
		env.node.scrollTop = 80;

		env.slideId.value = 's2';
		await nextTick();

		expect(env.store.getState().zoom).toBe(0.5);
		expect([env.node.scrollLeft, env.node.scrollTop]).toStrictEqual([0, 0]);
	});

	it('detaches when the component unmounts', async () => {
		env = await setup();
		const { node, store, wrapper } = env;
		wrapper.unmount();
		env = null;
		const event = ctrlWheel();
		node.dispatchEvent(event);
		expect(event.defaultPrevented).toBeFalsy();
		expect(store.getState().zoom).toBe(1);
	});
});
