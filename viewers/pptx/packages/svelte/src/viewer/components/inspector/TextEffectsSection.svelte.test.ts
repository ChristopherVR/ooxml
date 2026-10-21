import { TEXT_WARP_PRESETS } from 'ooxml-ui/pptx';
import type { PptxUiRibbonGalleryElement } from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { translationsDe } from 'pptx-viewer-locales';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { I18N_CONTEXT_KEY } from '../../../i18n/context';
import { interpolate, translate } from '../../../i18n/translator';
import type { Translator } from '../../../i18n/translator';
import { EditorState } from '../../editor/editor-state.svelte';
import TextEffectsSection from './TextEffectsSection.svelte';

/**
 * The Transform picker is the shared Text Warp gallery (`pptx-ui-ribbon-gallery`, a compact
 * panel). It used to be a select of ten short tokens (`archUp`) that are not OOXML presets:
 * the renderer and the saver know only the `text`-prefixed values (`textArchUp`), so a pick
 * stored a value nothing drew or saved, and a loaded deck's preset matched no option.
 */

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function textElement(textStyle: TextStyle = {}): PptxElement {
	return {
		type: 'text',
		id: 't1',
		x: 0,
		y: 0,
		width: 200,
		height: 50,
		content: 'Hello',
		textStyle: { fontSize: 24, ...textStyle },
	} as PptxElement;
}

interface Harness {
	editor: EditorState;
	gallery: PptxUiRibbonGalleryElement;
	/** Hand the edited element back to the section, as the inspector does on a selection change. */
	refresh(): void;
}

function mountSection(element: PptxElement, editable = true, translator?: Translator): Harness {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [element] }]);
	editor.select(element.id);
	editor.editable = editable;
	const target = document.createElement('div');
	document.body.appendChild(target);
	const props = $state({ editor, el: element });
	const instance = mount(TextEffectsSection, {
		target,
		props,
		...(translator ? { context: new Map([[I18N_CONTEXT_KEY, translator]]) } : {}),
	});
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	flushSync();
	// Found by its descriptor, the way the browser specs find it.
	const gallery = Array.from(target.querySelectorAll('pptx-ui-ribbon-gallery')).find(
		(candidate) => candidate.descriptor?.id === 'textWarp',
	);
	if (!gallery) {
		throw new Error('no Text Warp gallery');
	}
	return {
		editor,
		gallery,
		refresh: () => {
			props.el = editor.selectedElement ?? element;
			flushSync();
		},
	};
}

function tiles(gallery: HTMLElement): HTMLButtonElement[] {
	return Array.from(
		gallery.querySelectorAll<HTMLButtonElement>('.gallery-panel button.tile[data-gallery-item]'),
	);
}

function tile(gallery: HTMLElement, preset: string): HTMLButtonElement {
	const found = tiles(gallery).find((candidate) => candidate.dataset.galleryItem === preset);
	if (!found) {
		throw new Error(`no ${preset} tile`);
	}
	return found;
}

function pressed(gallery: HTMLElement): string[] {
	return tiles(gallery)
		.filter((candidate) => candidate.getAttribute('aria-pressed') === 'true')
		.map((candidate) => candidate.dataset.galleryItem ?? '');
}

function textStyleOf(editor: EditorState): TextStyle | undefined {
	const selected = editor.selectedElement;
	return selected && 'textStyle' in selected ? selected.textStyle : undefined;
}

describe('textEffectsSection text warp gallery', () => {
	it('re-labels its tiles when the viewer language changes', () => {
		// The viewer's translator reads its locale state on every call; this one does the same over
		// the German dictionary, without registering it for the other files of this worker.
		const locale = $state({ current: 'en' });
		const german = translationsDe as Record<string, string>;
		const t: Translator = (key, params) =>
			locale.current === 'de' && german[key]
				? interpolate(german[key], params)
				: translate('en', key, params);
		const { gallery } = mountSection(textElement(), true, t);
		expect(tile(gallery, 'textArchUp').getAttribute('aria-label')).toBe('Arch Up');
		locale.current = 'de';
		flushSync();
		expect(tile(gallery, 'textArchUp').getAttribute('aria-label')).toBe('Bogen aufwärts');
	});

	it('offers every OOXML preset as a compact tile, in the DOM while the section is closed', () => {
		const { gallery } = mountSection(textElement());

		expect(gallery.getAttribute('mode')).toBe('panel');
		expect(gallery.hasAttribute('compact')).toBeTruthy();
		expect(gallery.closest('details')?.open).toBeFalsy();
		expect(gallery.querySelector('.gallery-panel[role="group"]')?.getAttribute('aria-label')).toBe(
			'Text Warp',
		);
		const all = tiles(gallery);
		expect(all).toHaveLength(39);
		expect(all.map((candidate) => candidate.dataset.galleryItem)).toStrictEqual(
			TEXT_WARP_PRESETS.map((preset) => preset.value),
		);
		// Unlabelled tiles: the translated name is the tooltip and the accessible name.
		expect(tile(gallery, 'textArchUp').title).toBe('Arch Up');
		expect(tile(gallery, 'textArchUp').getAttribute('aria-label')).toBe('Arch Up');
		expect(gallery.querySelector('.tile-label')).toBeNull();
	});

	it('marks No Transform for unwarped text and a loaded preset when there is one', () => {
		expect(pressed(mountSection(textElement()).gallery)).toStrictEqual(['textNoShape']);
		cleanup?.();
		const loaded = mountSection(textElement({ textWarpPreset: 'textArchUp' }));
		expect(pressed(loaded.gallery)).toStrictEqual(['textArchUp']);
	});

	it('merges the picked preset into the text style and marks it', () => {
		const { editor, gallery, refresh } = mountSection(
			textElement({ text3d: { presetMaterial: 'metal' } }),
		);
		tile(gallery, 'textSlantUp').click();
		flushSync();

		expect(textStyleOf(editor)).toStrictEqual({
			fontSize: 24,
			text3d: { presetMaterial: 'metal' },
			textWarpPreset: 'textSlantUp',
		});
		refresh();
		expect(pressed(gallery)).toStrictEqual(['textSlantUp']);
	});

	it('clears the preset for No Transform', () => {
		const { editor, gallery, refresh } = mountSection(
			textElement({ textWarpPreset: 'textWave1', textWarpAdj: 30000 }),
		);
		tile(gallery, 'textNoShape').click();
		flushSync();

		const style = textStyleOf(editor);
		expect(style?.textWarpPreset).toBeUndefined();
		expect(style?.fontSize).toBe(24);
		refresh();
		expect(pressed(gallery)).toStrictEqual(['textNoShape']);
	});

	it('is disabled while the deck is not editable', () => {
		const { editor, gallery } = mountSection(textElement(), false);

		expect(gallery.disabled).toBeTruthy();
		expect(tiles(gallery).every((candidate) => candidate.disabled)).toBeTruthy();
		tile(gallery, 'textArchUp').click();
		flushSync();
		expect(textStyleOf(editor)?.textWarpPreset).toBeUndefined();

		editor.editable = true;
		flushSync();
		expect(tiles(gallery).some((candidate) => candidate.disabled)).toBeFalsy();
	});
});
