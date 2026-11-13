// @vitest-own-worker: mounts a full viewer.
import { ELEMENT_RENDER_FAULT_KEY } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { PptxHandler } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import type { BuildingBlocksContractHarness } from '../../../shared/test-utils/building-blocks-contract.mjs';
import { runElementRenderErrorContract } from '../../../shared/test-utils/element-render-error-contract.mjs';
import { PptxViewer } from '../viewer/PptxViewer';
import { createViewerBuildingBlocks } from './create-viewer-building-blocks';
import type { ViewerBuildingBlocks } from './types';

const tick = () =>
	new Promise<void>((resolve) => {
		setTimeout(resolve, 0);
	});

function vanillaHarness(): BuildingBlocksContractHarness {
	let blocks: ViewerBuildingBlocks | undefined;
	return {
		name: 'vanilla',
		toolbar: true,
		async mount(target, options) {
			blocks = createViewerBuildingBlocks(options);
			const pane = document.createElement('div');
			const canvas = document.createElement('div');
			target.append(pane, canvas);
			blocks.mountSlidesPane(pane);
			blocks.mountCanvas(canvas);
			await tick();
		},
		async update(options) {
			blocks!.update(options);
			await tick();
		},
		actions: () => blocks!.actions,
		state: () => blocks!.getState(),
		clickPictures: tick,
		settle: tick,
		destroy: () => blocks?.destroy(),
	};
}

const shape = (id: string, x: number) => ({
	id,
	type: 'shape' as const,
	x,
	y: 0,
	width: 50,
	height: 50,
	shapeType: 'rect',
	shapeStyle: { fillColor: '#336699' },
});

let viewer: PptxViewer | undefined;
afterEach(() => {
	viewer?.destroy();
	viewer = undefined;
	Reflect.deleteProperty(globalThis, Symbol.for(ELEMENT_RENDER_FAULT_KEY));
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

describe('vanilla element render errors', () => {
	it('passes the shared element render error contract', async () => {
		await runElementRenderErrorContract({
			harness: vanillaHarness(),
			expect,
			vi,
			Handler: PptxHandler,
		});
	}, 60000);

	it('omits only the failed element on the canvas and the rail, and retries a replacement', () => {
		const report = vi.fn();
		const container = document.createElement('div');
		document.body.append(container);
		viewer = new PptxViewer(container, { onElementRenderError: report });
		(globalThis as Record<symbol, unknown>)[Symbol.for(ELEMENT_RENDER_FAULT_KEY)] = 'bad';
		const bad = shape('bad', 60);
		const slide = (elements: PptxSlide['elements']): PptxSlide => ({
			id: 's1',
			rId: 'r1',
			slideNumber: 1,
			elements,
		});
		viewer.store.set({ slides: [slide([shape('good', 0), bad])] });
		const canvasIds = () =>
			[...container.querySelectorAll('.pptxv-stage [data-element-id]')].map((node) =>
				node.getAttribute('data-element-id'),
			);
		expect(canvasIds()).toContain('good');
		expect(canvasIds()).not.toContain('bad');
		// Canvas and rail each report once.
		const reported = report.mock.calls.length;
		expect(reported).toBeGreaterThanOrEqual(1);
		expect(report).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's1',
		});
		// Rebuilding with the same failed object neither retries nor reports again.
		viewer.store.set({ canvasSize: { width: 961, height: 540 } });
		expect(report).toHaveBeenCalledTimes(reported);
		// A replacement object is tried again.
		Reflect.deleteProperty(globalThis, Symbol.for(ELEMENT_RENDER_FAULT_KEY));
		viewer.store.set({ slides: [slide([shape('good', 0), { ...bad }])] });
		expect(canvasIds()).toContain('bad');
	});
});
