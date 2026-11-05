// @vitest-environment happy-dom
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { StaticElementRenderer } from './StaticElementRenderer';

/**
 * A group child's strokes snap to the device grid by their ABSOLUTE position in
 * the stage (the group origin plus the child's own offset), the way PowerPoint's
 * export places them, and a group child connector paints the same quantised
 * width as a top-level one (shared `device-pixel-stroke`).
 */
let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
});

function group(): PptxElement {
	return {
		type: 'group',
		id: 'g1',
		x: 200,
		y: 100,
		width: 300,
		height: 100,
		children: [
			{
				type: 'connector',
				id: 'c1',
				x: 10,
				y: 23.4,
				width: 50,
				height: 0,
				shapeStyle: { strokeColor: '#000000', strokeWidth: 0.3333 },
			},
		],
	} as unknown as PptxElement;
}

describe('react: group child strokes', () => {
	it('snaps a group child connector by its absolute y and quantises its width', () => {
		act(() => {
			root.render(
				<StaticElementRenderer
					element={group()}
					activeSlide={undefined as never}
					allSlides={[]}
					mediaDataUrls={new Map()}
					sourceSlideIndex={0}
					zIndex={0}
					exposeElementId
				/>,
			);
		});
		const line = container.querySelector<SVGPathElement>(
			'[data-element-id="c1"] path:last-of-type',
		);
		const style = line?.getAttribute('style') ?? '';
		// 100 (group) + 23.4 (child): the stroke snaps at 123.4, not at 23.4.
		expect(style).toContain('round(123.4px');
		expect(style).not.toContain('round(23.4px');
		// The width is a whole number of device pixels, not `max(0.3333px, ...)`.
		expect(style).toContain('round(0.3333px');
	});
});
