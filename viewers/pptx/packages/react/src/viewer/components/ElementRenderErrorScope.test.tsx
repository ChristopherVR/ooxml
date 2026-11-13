// @vitest-environment happy-dom
/**
 * `ElementRenderErrorScope` (what `PowerPointViewer`'s `onElementRenderError` provides) reaches
 * every element boundary below it unless a nearer surface passes its own reporter, and the shared
 * test fault hook throws from inside the boundary.
 */
import { ELEMENT_RENDER_FAULT_KEY } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { ElementRenderErrorHandler } from './ElementErrorBoundary';
import { ElementErrorBoundary, ElementRenderErrorScope } from './ElementErrorBoundary';

const FAULT = Symbol.for(ELEMENT_RENDER_FAULT_KEY);
const element: PptxElement = {
	id: 'bad',
	type: 'shape',
	x: 0,
	y: 0,
	width: 10,
	height: 10,
	shapeType: 'rect',
};

let container: HTMLDivElement;
let root: Root;
beforeEach(() => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	vi.spyOn(console, 'error').mockImplementation(() => {});
	container = document.createElement('div');
	document.body.append(container);
	root = createRoot(container);
});
afterEach(() => {
	act(() => root.unmount());
	container.remove();
	Reflect.deleteProperty(globalThis, FAULT);
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

function Body(): React.ReactElement {
	throw new Error('body failed');
}

describe('ElementRenderErrorScope', () => {
	it('gives boundaries below it the viewer-level reporter', () => {
		const viewerReport = vi.fn<ElementRenderErrorHandler>();
		act(() =>
			root.render(
				<ElementRenderErrorScope onElementRenderError={viewerReport}>
					<ElementErrorBoundary element={element} slideId='s1'>
						<Body />
					</ElementErrorBoundary>
					<span data-healthy>healthy</span>
				</ElementRenderErrorScope>,
			),
		);
		expect(container.querySelector('[data-healthy]')).not.toBeNull();
		expect(viewerReport).toHaveBeenCalledWith(expect.any(Error), {
			elementId: 'bad',
			elementType: 'shape',
			slideId: 's1',
		});
	});

	it('lets a nearer reporter win', () => {
		const viewerReport = vi.fn<ElementRenderErrorHandler>();
		const own = vi.fn<ElementRenderErrorHandler>();
		act(() =>
			root.render(
				<ElementRenderErrorScope onElementRenderError={viewerReport}>
					<ElementErrorBoundary element={element} onElementRenderError={own}>
						<Body />
					</ElementErrorBoundary>
				</ElementRenderErrorScope>,
			),
		);
		expect(own).toHaveBeenCalledTimes(1);
		expect(viewerReport).not.toHaveBeenCalled();
	});
});
