/**
 * Which `SlideStage`s keep `data-element-id`: the editing canvas, the show
 * stage and Morph layers do; every static stage (rail, host thumbnails,
 * sorter, previews) drops it, so element ids stay unique to the canvas.
 */
import type { Rendering3DFlags } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import {
	DEFAULT_RENDERING_3D_FLAGS,
	Rendering3DFlagsContextKey,
} from '../state/rendering-3d-flags-context';
import type { SlideStageProps } from './props';
import SlideStage from './SlideStage.svelte';

const canvasSize = { width: 960, height: 540 };

const shape = (id: string): PptxElement =>
	({ id, type: 'shape', shapeType: 'rect', x: 0, y: 0, width: 100, height: 50 }) as PptxElement;

const smartArt = {
	id: 'sa-1',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	smartArtData: {
		layoutType: 'list',
		nodes: [{ id: 'n1', text: 'One' }],
		drawingShapes: [
			{ id: 'a', shapeType: 'roundRect', x: 0, y: 0, width: 400, height: 140, text: 'One' },
		],
	},
} as unknown as PptxElement;

const slideWith = (elements: PptxElement[]): PptxSlide =>
	({ id: 's1', rId: 'rId1', slideNumber: 1, elements }) as PptxSlide;

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

/** Mount a stage; the strip runs one microtask after each render. */
function mountStage(
	extra: Partial<SlideStageProps>,
	context = new Map<unknown, unknown>(),
): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(SlideStage, {
		target,
		props: {
			slide: slideWith([shape('shape-1')]),
			canvasSize,
			mediaDataUrls: new Map<string, string>(),
			...extra,
		},
		context,
	});
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

describe('slideStage element-id markers', () => {
	it('keeps them on the editing canvas, the show stage and Morph layers', async () => {
		for (const flags of [
			{ interactive: true },
			{ presenting: true },
			{ preserveElementIds: true },
		]) {
			const target = mountStage(flags);
			await Promise.resolve();
			expect(target.querySelector('[data-element-id="shape-1"]')).not.toBeNull();
			cleanup?.();
			cleanup = undefined;
		}
	});

	it('drops them from a static stage, which still draws the element', async () => {
		const target = mountStage({});
		await Promise.resolve();
		expect(target.querySelector('[data-element-id]')).toBeNull();
		expect(target.querySelectorAll('.pptx-svelte-element')).toHaveLength(1);
	});

	it('drops them again when a 3D option swaps an element renderer', async () => {
		// Turning SmartArt 3D on mounts a new renderer node for the same element
		// list, so only a dependency on the options re-runs the strip.
		const flags = $state<Rendering3DFlags>({ ...DEFAULT_RENDERING_3D_FLAGS, smartArt3D: false });
		const target = mountStage(
			{ slide: slideWith([smartArt]) },
			new Map([[Rendering3DFlagsContextKey, () => flags]]),
		);
		await Promise.resolve();
		expect(target.querySelector('pptx-three-view')).toBeNull();
		expect(target.querySelector('[data-element-id]')).toBeNull();
		flags.smartArt3D = true;
		flushSync();
		await Promise.resolve();
		expect(target.querySelector('pptx-three-view')).not.toBeNull();
		expect(target.querySelector('[data-element-id]')).toBeNull();
	});
});
