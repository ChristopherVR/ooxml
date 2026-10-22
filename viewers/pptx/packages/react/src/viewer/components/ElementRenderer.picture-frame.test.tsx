import { translationsEn } from 'ooxml-ui/pptx/i18n';
// @vitest-environment happy-dom
/**
 * A picture's own frame (`a:ln`) is centred on the picture edge, so its inner
 * half lies over the photo. React stroked it in `ShapeEffectOverlay`, but that
 * overlay was mounted before the picture surface, so the photo painted over
 * the inner half and only the outer half of the line showed. Both React
 * renderers must paint a picture's overlay after the picture.
 */
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ElementRenderer } from './ElementRenderer';
import type { ElementRendererProps } from './elements/element-renderer-types';
import { StaticElementRenderer } from './StaticElementRenderer';

vi.mock<typeof import('../i18n/use-translation')>(import('../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => translationsEn[key] ?? key }),
}));

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

const PICTURE = {
	type: 'picture',
	id: 'pic-frame',
	x: 0,
	y: 0,
	width: 200,
	height: 150,
	imageData: 'data:image/png;base64,AA==',
	shapeStyle: { strokeWidth: 10.67, strokeFillMode: 'solid', strokeColor: '#FF0000' },
} as unknown as PptxElement;

function props(element: PptxElement): ElementRendererProps {
	return {
		element,
		isSelected: false,
		isInlineEditing: false,
		inlineEditingText: '',
		canInteract: false,
		spellCheckEnabled: false,
		mediaDataUrls: new Map(),
		selectionColorClass: 'blue-500',
		showHoverBorder: false,
		imageAltText: 'Slide element',
		showResizeHandles: false,
		renderInk: true,
		renderGroups: true,
		adjustmentHandles: [],
		onResizePointerDown: vi.fn<() => void>(),
		onAdjustmentPointerDown: vi.fn<() => void>(),
		onInlineEditChange: vi.fn<() => void>(),
		onInlineEditCommit: vi.fn<() => void>(),
		onInlineEditCancel: vi.fn<() => void>(),
	};
}

/** True when the red frame stroke comes after the picture's img in paint (tree) order. */
function frameFollowsImage(): boolean {
	const img = container.querySelector('img');
	const frame = [...container.querySelectorAll('svg path')].find(
		(p) => p.getAttribute('stroke') === '#FF0000',
	);
	expect(img).not.toBeNull();
	expect(frame).toBeDefined();
	return Boolean(img!.compareDocumentPosition(frame!) & Node.DOCUMENT_POSITION_FOLLOWING);
}

describe('react picture frame paints over the picture', () => {
	it('keeps its separately animated background frame above the image', () => {
		act(() =>
			root.render(
				<ElementRenderer
					{...props(PICTURE)}
					presentationElementStates={
						new Map([
							[
								`${PICTURE.id}::pptx-bg`,
								{ visible: true, cssAnimation: 'pptx-fade-in 500ms ease 0ms 1 normal both' },
							],
						])
					}
				/>,
			),
		);
		expect(frameFollowsImage()).toBeTruthy();
		const background = container.querySelector<HTMLElement>(
			'[data-pptx-animation-layer="background"]',
		);
		expect(background?.style.animation).toContain('pptx-fade-in');
		expect(background?.querySelector('svg path')).not.toBeNull();
	});

	it('on the interactive canvas', () => {
		act(() => root.render(<ElementRenderer {...props(PICTURE)} />));
		expect(frameFollowsImage()).toBe(true);
	});

	it('on static surfaces (thumbnails, previews, export)', () => {
		act(() => root.render(<StaticElementRenderer element={PICTURE} positioned />));
		expect(frameFollowsImage()).toBe(true);
	});
});
