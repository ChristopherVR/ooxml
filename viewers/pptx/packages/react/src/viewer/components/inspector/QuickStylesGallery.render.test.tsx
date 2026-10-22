// @vitest-environment happy-dom
/**
 * Quick Styles draws its presets on the shared gallery (`pptx-ui-ribbon-gallery`, a compact
 * panel) and hands a pick's shape style to the inspector, which merges it over the shape's own.
 */
import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { QuickStylesGallery } from './QuickStylesGallery';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

const tiles = () => [...container.querySelectorAll<HTMLButtonElement>('[data-gallery-item]')];

describe('quickStylesGallery', () => {
	it('draws every preset on the shared compact panel', () => {
		act(() => root.render(<QuickStylesGallery onUpdateShapeStyle={() => {}} />));
		const gallery = container.querySelector('pptx-ui-ribbon-gallery')!;
		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(tiles().map((tile) => tile.getAttribute('aria-label'))).toStrictEqual(
			SHAPE_QUICK_STYLES.map((preset) => preset.name),
		);
		expect(tiles()[0]!.querySelector('svg')).not.toBeNull();
	});

	it("hands the picked preset's shape style to the inspector", () => {
		const update = vi.fn();
		act(() => root.render(<QuickStylesGallery onUpdateShapeStyle={update} />));
		act(() => tiles()[6]!.click());
		expect(update.mock.calls).toStrictEqual([[SHAPE_QUICK_STYLES[6]!.style]]);
	});
});
