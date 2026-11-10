// @vitest-environment happy-dom
/**
 * The vanilla wiring of the shared viewport navigation: Ctrl+wheel zooms in the
 * same fit-relative steps as the other bindings (the store keeps an absolute
 * scale), a plain drag pans only in a read-only view, and destroy detaches.
 */
import { attachViewportNavigation } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PptxViewer } from './PptxViewer';

const slide: PptxSlide = { id: 'slide', rId: 'rId1', slideNumber: 1, elements: [] };

let viewer: PptxViewer | undefined;
afterEach(() => {
	viewer?.destroy();
	viewer = undefined;
	document.body.replaceChildren();
	vi.restoreAllMocks();
});

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

/** The fit-relative factor one wheel notch moves to from fit, per the shared controller. */
function sharedWheelStep(): number {
	const probe = document.createElement('div');
	let zoom = 1;
	const detach = attachViewportNavigation(probe, {
		getZoom: () => zoom,
		setZoom: (next) => (zoom = next),
	});
	probe.dispatchEvent(ctrlWheel());
	detach();
	return zoom;
}

function mount(editable: boolean) {
	const client = { width: 1008, height: 588 };
	for (const [key, value] of [
		['clientWidth', () => client.width],
		['clientHeight', () => client.height],
	] as const) {
		vi.spyOn(HTMLElement.prototype, key, 'get').mockImplementation(function (this: HTMLElement) {
			return this.hasAttribute('data-pptx-viewport') ? value() : 0;
		});
	}
	const container = document.createElement('div');
	document.body.append(container);
	viewer = new PptxViewer(container, { fitPadding: 0, maxFitScale: null });
	viewer.store.set({ slides: [slide], canvasSize: { width: 960, height: 540 }, editable });
	const viewport = container.querySelector<HTMLElement>('[data-pptx-viewport]')!;
	let scroll = 0;
	Object.defineProperties(viewport, {
		scrollWidth: { configurable: true, get: () => 3000 },
		scrollHeight: { configurable: true, get: () => 3000 },
		scrollLeft: {
			configurable: true,
			get: () => scroll,
			set: (value: number) => {
				scroll = Math.max(0, value);
			},
		},
		scrollTop: { configurable: true, get: () => 0, set: () => {} },
	});
	vi.spyOn(viewport, 'setPointerCapture').mockImplementation(() => {});
	vi.spyOn(viewport, 'releasePointerCapture').mockImplementation(() => {});
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
			viewport.dispatchEvent(event);
		};
		pointer('pointerdown', 200);
		pointer('pointermove', 150);
		pointer('pointerup', 150);
		return scroll;
	};
	return { viewer, viewport, drag };
}

describe('vanilla viewport navigation', () => {
	it('zooms on Ctrl+wheel by the shared fit-relative step', () => {
		const env = mount(true);
		const event = ctrlWheel();
		env.viewport.dispatchEvent(event);
		expect(event.defaultPrevented).toBeTruthy();
		// 1008 / 960 = 1.05 fits the slide; the store holds the absolute scale.
		expect(env.viewer.store.get().zoom).toBeCloseTo(1.05 * sharedWheelStep());
	});

	it('pans with a plain drag only in a read-only view', () => {
		const env = mount(false);
		expect(env.drag()).toBe(150);
		env.viewer.store.set({ editable: true });
		expect(env.drag()).toBe(100);
	});

	it('detaches on destroy', () => {
		const env = mount(true);
		env.viewer.destroy();
		viewer = undefined;
		const event = ctrlWheel();
		env.viewport.dispatchEvent(event);
		expect(event.defaultPrevented).toBeFalsy();
	});
});
