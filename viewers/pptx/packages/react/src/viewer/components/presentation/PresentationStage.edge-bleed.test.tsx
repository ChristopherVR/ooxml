// @vitest-environment happy-dom
/**
 * Issue #141: two abutting filled rectangles must not show a light seam where
 * their shared edge falls inside a device pixel. The lower box grows under the
 * opaque upper one (shared `resolveStageEdgeBleeds`), on the show stage as on
 * the editing canvas; the upper box and every other element keep their box.
 */
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { PresentationStage } from './PresentationStage';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function rect(id: string, x: number, fillColor: string): PptxElement {
	return {
		type: 'shape',
		id,
		x,
		y: 0,
		width: 100,
		height: 100,
		shapeType: 'rect',
		text: '',
		shapeStyle: { fillMode: 'solid', fillColor, strokeWidth: 0, strokeFillMode: 'none' },
	} as PptxElement;
}

function render(slide: PptxSlide): void {
	act(() =>
		root.render(
			<PresentationStage
				activeSlide={slide}
				templateElements={[]}
				canvasSize={{ width: 960, height: 540 }}
				mediaDataUrls={new Map<string, string>()}
				onStageClick={() => undefined}
			/>,
		),
	);
}

const box = (id: string): HTMLElement =>
	container.querySelector<HTMLElement>(`[data-element-id="${id}"]`)!;

describe('presentationStage abutting-edge bleed', () => {
	it('grows the lower box under the abutting opaque box above it', () => {
		render({
			id: 's',
			elements: [rect('lower', 100, '#112233'), rect('upper', 200, '#445566')],
		} as unknown as PptxSlide);
		expect(box('lower').style.left).toBe('100px');
		expect(box('lower').style.width).toBe('102px');
		expect(box('upper').style.width).toBe('100px');
	});

	it('keeps exact boxes when the slide animates the upper box', () => {
		render({
			id: 's',
			elements: [rect('lower', 100, '#112233'), rect('upper', 200, '#445566')],
			animations: [{ elementId: 'upper' }],
		} as unknown as PptxSlide);
		expect(box('lower').style.width).toBe('100px');
	});
});
