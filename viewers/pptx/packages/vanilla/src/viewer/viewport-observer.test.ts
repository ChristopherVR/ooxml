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
	vi.unstubAllGlobals();
});

interface Observation {
	box: string | undefined;
	target: Element;
	callback: () => void;
	disconnected: boolean;
}

/** A ResizeObserver stub that records every observation. */
function stubObservers(): Observation[] {
	const observations: Observation[] = [];
	vi.stubGlobal(
		'ResizeObserver',
		class {
			private readonly mine: Observation[] = [];
			private readonly callback: () => void;
			constructor(callback: () => void) {
				this.callback = callback;
			}
			observe(target: Element, options?: ResizeObserverOptions) {
				const observation = {
					box: options?.box,
					target,
					callback: this.callback,
					disconnected: false,
				};
				this.mine.push(observation);
				observations.push(observation);
			}
			unobserve() {}
			disconnect() {
				for (const observation of this.mine) {
					observation.disconnected = true;
				}
			}
		},
	);
	return observations;
}

describe('vanilla viewport observation', () => {
	it('keeps lifecycle.resizeObserver on the content box and disconnects it on destroy', () => {
		const observations = stubObservers();
		const container = document.createElement('div');
		document.body.append(container);
		viewer = new PptxViewer(container);
		const viewport = container.querySelector('[data-pptx-viewport]')!;
		const resizeObserver = viewer.lifecycle?.resizeObserver;
		expect(resizeObserver).toBeTruthy();
		const contentBox = observations.find(
			(observation) => observation.target === viewport && observation.box === undefined,
		);
		expect(contentBox).toBeDefined();
		resizeObserver!.disconnect();
		expect(contentBox!.disconnected).toBeTruthy();
	});

	it('refits on a padding-only change, which only the border box shows', () => {
		const observations = stubObservers();
		// A content-box viewport losing its padding: the client size drops, the content box a
		// default ResizeObserver watches stays put.
		const client = { width: 1008, height: 588 };
		vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockImplementation(function (
			this: HTMLElement,
		) {
			return this.hasAttribute('data-pptx-viewport') ? client.width : 0;
		});
		vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockImplementation(function (
			this: HTMLElement,
		) {
			return this.hasAttribute('data-pptx-viewport') ? client.height : 0;
		});
		const container = document.createElement('div');
		document.body.append(container);
		viewer = new PptxViewer(container, { fitPadding: 0, maxFitScale: null });
		viewer.store.set({ slides: [slide], canvasSize: { width: 960, height: 540 } });
		const viewport = container.querySelector('[data-pptx-viewport]')!;
		const stage = () => viewport.querySelector<HTMLElement>('[aria-roledescription="slide"]')!;
		expect(stage().style.transform).toBe('scale(1.05)');
		const borderBox = observations.find(
			(observation) => observation.target === viewport && observation.box === 'border-box',
		);
		expect(borderBox).toBeDefined();
		client.width = 960;
		client.height = 540;
		borderBox!.callback();
		expect(stage().style.transform).toBe('scale(1)');
		viewer.destroy();
		viewer = undefined;
		expect(
			observations
				.filter((observation) => observation.target === viewport)
				.every((observation) => observation.disconnected),
		).toBeTruthy();
	});
});
