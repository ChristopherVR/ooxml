/**
 * `ElementBoundary` (a `<svelte:boundary>` per element): a failing element is omitted alone and
 * reported; the same failed object is not retried, a replacement object or a new id is.
 */
import { ELEMENT_RENDER_FAULT_KEY } from 'ooxml-ui/pptx';
import type { ElementRenderErrorHandler } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, tick, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { ElementRenderErrorScopeKey } from '../state/element-render-errors';
import type { ElementRenderErrorScopeGetter } from '../state/element-render-errors';
import SlideStage from './SlideStage.svelte';

const FAULT = Symbol.for(ELEMENT_RENDER_FAULT_KEY);
const setFault = (id: string | undefined) => {
	(globalThis as Record<symbol, unknown>)[FAULT] = id;
};
const shape = (id: string, x = 0): PptxElement =>
	({ id, type: 'shape', shapeType: 'rect', x, y: 0, width: 100, height: 50 }) as PptxElement;
const slideWith = (elements: PptxElement[]): PptxSlide =>
	({ id: 's1', rId: 'rId1', slideNumber: 1, elements }) as PptxSlide;

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	Reflect.deleteProperty(globalThis, FAULT);
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

class SlideHolder {
	slide = $state.raw<PptxSlide>({} as PptxSlide);
	constructor(initial: PptxSlide) {
		this.slide = initial;
	}
}

async function settle(): Promise<void> {
	flushSync();
	await tick();
	flushSync();
}

function mountStage(initial: PptxSlide, report: ElementRenderErrorHandler | undefined) {
	const target = document.createElement('div');
	document.body.appendChild(target);
	// Raw state: the test controls element identity, as the viewer's deck state does.
	const holder = new SlideHolder(initial);
	const props = {
		get slide() {
			return holder.slide;
		},
		set slide(next: PptxSlide) {
			holder.slide = next;
		},
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map<string, string>(),
		interactive: true,
	};
	const scope: ElementRenderErrorScopeGetter = () => ({ onElementRenderError: report });
	const instance = mount(SlideStage, {
		target,
		props,
		context: new Map<unknown, unknown>([[ElementRenderErrorScopeKey, scope]]),
	});
	cleanup = () => void unmount(instance);
	const ids = () =>
		[...target.querySelectorAll('[data-element-id]')].map((node) =>
			node.getAttribute('data-element-id'),
		);
	return { props, ids };
}

describe('Svelte ElementBoundary', () => {
	it('omits only the failing element, then retries a replacement object or a new id', async () => {
		setFault('bad');
		const report = vi.fn<ElementRenderErrorHandler>();
		const bad = shape('bad', 120);
		const { props, ids } = mountStage(slideWith([shape('good'), bad]), report);
		await settle();
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenCalledTimes(1);
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's1',
		});
		// The same failed object is not retried when the slide re-renders.
		props.slide = slideWith([shape('good'), bad]);
		await settle();
		expect(report).toHaveBeenCalledTimes(1);
		// A replacement object is retried (and fails again while the fault holds).
		props.slide = slideWith([shape('good'), { ...bad }]);
		await settle();
		expect(report).toHaveBeenCalledTimes(2);
		// A healthy replacement renders.
		setFault(undefined);
		props.slide = slideWith([shape('good'), { ...bad }]);
		await settle();
		expect(ids()).toEqual(['good', 'bad']);
		// A changed id gets a boundary of its own.
		setFault('renamed');
		props.slide = slideWith([shape('good'), shape('renamed', 120)]);
		await settle();
		expect(ids()).toEqual(['good']);
		expect(report).toHaveBeenLastCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'renamed' }),
		);
	});

	it('falls back to console.error when the reporter throws', async () => {
		setFault('bad');
		const log = vi.spyOn(console, 'error').mockImplementation(() => {});
		const throwing = vi.fn(() => {
			throw new Error('reporter');
		});
		const { ids } = mountStage(slideWith([shape('good'), shape('bad', 120)]), throwing);
		await settle();
		expect(throwing).toHaveBeenCalledTimes(1);
		expect(ids()).toEqual(['good']);
		expect(log).toHaveBeenCalledWith(
			expect.any(Error),
			expect.objectContaining({ elementId: 'bad' }),
		);
	});
});
