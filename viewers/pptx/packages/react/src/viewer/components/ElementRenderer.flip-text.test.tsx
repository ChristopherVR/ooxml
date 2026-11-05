// @vitest-environment happy-dom
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
/**
 * Text in a flipped shape, rendered THROUGH `ElementRenderer`.
 *
 * PowerPoint 16 never mirrors text: `flipH` paints it as the unflipped shape
 * does and `flipV` paints it upside down. Inside the container's
 * `scaleY(-1)` the text layer's `scaleX(-1)` makes that half turn, and it
 * pivots on the text rectangle's centre, which a lopsided inset moves off
 * the box centre.
 */
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';

import { ElementRenderer } from './ElementRenderer';

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

function box(overrides: Partial<PptxElement>): PptxElement {
	return {
		type: 'text',
		id: 't',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		text: 'Lorem',
		textSegments: [{ text: 'Lorem', style: {} }],
		textStyle: { fontSize: 18 },
		...overrides,
	} as PptxElement;
}

/** The transform and origin of the element holding the text's paragraphs. */
function textLayer(element: PptxElement): { transform: string; origin: string } {
	act(() => {
		root.render(
			<ElementRenderer element={element} zIndex={0} mediaDataUrls={new Map<string, string>()} />,
		);
	});
	const layers = [...container.querySelectorAll<HTMLElement>('*')].filter((node) =>
		node.style.transform.includes('scaleX(-1)'),
	);
	const layer = layers.find(
		(node) =>
			node.textContent === 'Lorem' && node.parentElement?.style.transform !== node.style.transform,
	);
	return {
		transform: layer?.style.transform ?? '',
		origin: layer?.style.transformOrigin ?? '',
	};
}

describe('elementRenderer - flipped text', () => {
	it('turns vertically flipped text upside down instead of mirroring it', () => {
		const { transform } = textLayer(box({ flipVertical: true }));
		expect(transform).toBe('scaleX(-1)');
	});

	it('pivots on the text rectangle when the insets differ', () => {
		const { origin } = textLayer(
			box({
				flipVertical: true,
				textStyle: { fontSize: 18, bodyInsetLeft: 40, bodyInsetRight: 0 },
			} as Partial<PptxElement>),
		);
		expect(origin).toBe('120px 50%');
	});
});
