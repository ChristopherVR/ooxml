import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import type { PptxUiRibbonGalleryElement } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import QuickStylesGallery from './QuickStylesGallery.svelte';

/**
 * Quick Styles draws its presets on the shared gallery (`pptx-ui-ribbon-gallery`, a compact
 * panel), so these read the tiles the element draws in its light DOM rather than any markup of
 * this binding.
 */

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function shapeEl(): PptxElement {
	return {
		type: 'shape',
		id: 's1',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		// `fillOpacity` is set by no preset, so its survival proves a merge, not a replace.
		shapeStyle: { strokeWidth: 7, fillOpacity: 0.42 },
	} as PptxElement;
}

interface Harness {
	target: HTMLElement;
	editor: EditorState;
	gallery: PptxUiRibbonGalleryElement;
}

function mountGallery(editable = true): Harness {
	const el = shapeEl();
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = editable;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [el] }]);
	editor.editable = editable;
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(QuickStylesGallery, { target, props: { editor, el } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	const gallery = target.querySelector('pptx-ui-ribbon-gallery');
	if (!gallery) {
		throw new Error('no Quick Styles gallery');
	}
	return { target, editor, gallery };
}

function tiles(gallery: HTMLElement): HTMLButtonElement[] {
	return Array.from(gallery.querySelectorAll<HTMLButtonElement>('[data-gallery-item]'));
}

function shapeStyleOf(editor: EditorState): ShapeStyle | undefined {
	return (editor.slides[0]?.elements?.[0] as { shapeStyle?: ShapeStyle } | undefined)?.shapeStyle;
}

describe('quickStylesGallery', () => {
	it('renders one accessibly named tile per shared preset on the compact panel', () => {
		const { target, gallery } = mountGallery();

		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(gallery.descriptor?.id).toBe('quickStyles');
		expect(gallery.querySelector('.gallery-panel[role="group"]')?.getAttribute('aria-label')).toBe(
			'Quick Styles',
		);
		const names = SHAPE_QUICK_STYLES.map((preset) => preset.name);
		expect(names).toHaveLength(18);
		expect(tiles(gallery).map((tile) => tile.getAttribute('aria-label'))).toStrictEqual(names);
		expect(tiles(gallery).map((tile) => tile.title)).toStrictEqual(names);
		// The caption is a plain span inside the wrapper the inspector looks for.
		const caption = target.querySelector('.pptx-svelte-quick-styles > span');
		expect(caption?.textContent).toBe('Quick Styles');
		expect(target.querySelector('label')).toBeNull();
	});

	it('merges the preset over the existing shape style rather than replacing it', () => {
		const { gallery, editor } = mountGallery();
		const preset = SHAPE_QUICK_STYLES[6]!;

		tiles(gallery)[6]?.click();
		flushSync();

		expect(shapeStyleOf(editor)).toStrictEqual({ fillOpacity: 0.42, ...preset.style });
	});

	it('disables every tile in a read-only viewer, and enables them once editable', () => {
		const { gallery, editor } = mountGallery(false);

		expect(gallery.disabled).toBeTruthy();
		expect(tiles(gallery)).toHaveLength(SHAPE_QUICK_STYLES.length);
		expect(tiles(gallery).every((tile) => tile.disabled)).toBeTruthy();
		tiles(gallery)[0]?.click();
		flushSync();
		expect(shapeStyleOf(editor)).toStrictEqual({ strokeWidth: 7, fillOpacity: 0.42 });

		editor.editable = true;
		flushSync();
		expect(tiles(gallery).some((tile) => tile.disabled)).toBeFalsy();
	});
});
