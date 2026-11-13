import { flushPromises, mount } from '@vue/test-utils';
import { ELEMENT_RENDER_FAULT_KEY } from 'ooxml-ui/pptx';
import type { ElementRenderErrorHandler } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { defineComponent, h, shallowRef } from 'vue';

import {
	isEventHandlerErrorInfo,
	provideElementRenderErrorReporter,
} from '../composables/element-render-errors';
import SlideStage from './SlideStage.vue';

const FAULT = Symbol.for(ELEMENT_RENDER_FAULT_KEY);
const setFault = (id: string | undefined) => {
	(globalThis as Record<symbol, unknown>)[FAULT] = id;
};

function shape(id: string, x = 0): PptxElement {
	return { type: 'shape', id, x, y: 0, width: 100, height: 40, text: id } as PptxElement;
}
const slideWith = (elements: PptxElement[]): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements }) as PptxSlide;

function mountStage(initial: PptxSlide, report: ElementRenderErrorHandler | undefined) {
	const slide = shallowRef(initial);
	const Host = defineComponent({
		setup() {
			provideElementRenderErrorReporter(() => report);
			return () =>
				h(SlideStage, {
					slide: slide.value,
					canvasSize: { width: 960, height: 540 },
					mediaDataUrls: new Map<string, string>(),
					interactive: true,
				});
		},
	});
	const wrapper = mount(Host, { attachTo: document.body });
	const ids = () =>
		[...wrapper.element.querySelectorAll('[data-element-id]')].map((node) =>
			node.getAttribute('data-element-id'),
		);
	return { wrapper, slide, ids };
}

afterEach(() => {
	Reflect.deleteProperty(globalThis, FAULT);
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

describe('Vue ElementBoundary', () => {
	it('omits only the failing element, then retries a replacement object or a new id', async () => {
		setFault('bad');
		const report = vi.fn<ElementRenderErrorHandler>();
		const bad = shape('bad', 120);
		const { wrapper, slide, ids } = mountStage(slideWith([shape('good'), bad]), report);
		await flushPromises();
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenCalledTimes(1);
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's1',
		});
		// The same failed object is not retried when the slide re-renders.
		slide.value = slideWith([shape('good'), bad]);
		await flushPromises();
		expect(report).toHaveBeenCalledTimes(1);
		// A replacement object is retried (and fails again while the fault holds).
		slide.value = slideWith([shape('good'), { ...bad }]);
		await flushPromises();
		expect(report).toHaveBeenCalledTimes(2);
		// A healthy replacement renders.
		setFault(undefined);
		slide.value = slideWith([shape('good'), { ...bad }]);
		await flushPromises();
		expect(ids()).toEqual(['good', 'bad']);
		// A changed id remounts the boundary.
		setFault('renamed');
		slide.value = slideWith([shape('good'), shape('renamed', 120)]);
		await flushPromises();
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenLastCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'renamed' }),
		);
		wrapper.unmount();
	});

	it('falls back to console.error when the reporter throws or is absent', async () => {
		setFault('bad');
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		const throwing = vi.fn(() => {
			throw new Error('reporter');
		});
		const first = mountStage(slideWith([shape('good'), shape('bad', 120)]), throwing);
		await flushPromises();
		expect(throwing).toHaveBeenCalledTimes(1);
		expect(first.ids()).toEqual(['good']);
		expect(log).toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'bad' }),
		);
		first.wrapper.unmount();
		log.mockClear();
		const second = mountStage(slideWith([shape('good'), shape('bad', 120)]), undefined);
		await flushPromises();
		expect(second.ids()).toEqual(['good']);
		expect(log).toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'bad' }),
		);
		second.wrapper.unmount();
	});

	it('leaves event handler errors to the app', () => {
		expect(isEventHandlerErrorInfo('native event handler')).toBe(true);
		expect(isEventHandlerErrorInfo('https://vuejs.org/error-reference/#runtime-6')).toBe(true);
		expect(isEventHandlerErrorInfo('render function')).toBe(false);
		expect(isEventHandlerErrorInfo('https://vuejs.org/error-reference/#runtime-1')).toBe(false);
	});
});
