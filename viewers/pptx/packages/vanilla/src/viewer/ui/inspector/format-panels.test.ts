import { SHAPE_QUICK_STYLES } from 'ooxml-ui/pptx';
import type { MediaPptxElement, TextStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createImageSection } from './image-section';
import { createMediaTrimTimeline } from './media-trim-timeline';
import { createQuickStylesGallery } from './quick-styles-gallery';
import { createText3DSection } from './text-3d-section';
import type { InspectorHandlers, InspectorState } from './types';

/** A `section()` factory matching the one `createInspector` passes in. */
function sectionFactory() {
	return (): HTMLElement => document.createElement('div');
}

function state(overrides: Partial<InspectorState> = {}): InspectorState {
	return { hasSelection: true, ...overrides } as InspectorState;
}

describe('quick styles gallery', () => {
	function build() {
		const setShapeStyle = vi.fn();
		const gallery = createQuickStylesGallery(document, createTranslator(), sectionFactory(), {
			setShapeStyle,
		});
		const tiles = () => [...gallery.el.querySelectorAll<HTMLButtonElement>('[data-gallery-item]')];
		return { gallery, setShapeStyle, tiles };
	}

	it('draws every preset in order on the shared compact gallery panel', () => {
		const { gallery, tiles } = build();
		gallery.update(state({ canShape: true }));

		const element = gallery.el.querySelector('pptx-ui-ribbon-gallery')!;
		expect(element.getAttribute('mode')).toBe('panel');
		expect(element.hasAttribute('compact')).toBeTruthy();
		expect(element.descriptor?.id).toBe('quickStyles');
		expect(tiles().map((tile) => tile.getAttribute('aria-label'))).toStrictEqual(
			SHAPE_QUICK_STYLES.map((preset) => preset.name),
		);
		expect(tiles()[0]!.querySelector('svg')).not.toBeNull();
	});

	it("hands the picked preset's style to setShapeStyle", () => {
		const { gallery, setShapeStyle, tiles } = build();
		gallery.update(state({ canShape: true }));

		tiles()[0]!.click();
		tiles()[12]!.click();

		expect(setShapeStyle.mock.calls).toStrictEqual([
			[SHAPE_QUICK_STYLES[0]!.style],
			[SHAPE_QUICK_STYLES[12]!.style],
		]);
	});

	it('keeps its tiles, and its picks, once the detached inspector is inserted', () => {
		const { gallery, setShapeStyle, tiles } = build();
		gallery.update(state({ canShape: true }));
		document.body.append(gallery.el);
		try {
			expect(tiles()).toHaveLength(SHAPE_QUICK_STYLES.length);
			tiles()[6]!.click();
			expect(setShapeStyle.mock.calls).toStrictEqual([[SHAPE_QUICK_STYLES[6]!.style]]);
		} finally {
			gallery.el.remove();
		}
	});

	it('hides and disables itself for a selection that cannot take a shape style', () => {
		const { gallery, setShapeStyle, tiles } = build();
		gallery.update(state({ canShape: false }));

		expect(gallery.el.hidden).toBeTruthy();
		expect(tiles()).toHaveLength(SHAPE_QUICK_STYLES.length);
		expect(tiles().every((tile) => tile.disabled)).toBeTruthy();
		tiles()[0]!.click();
		expect(setShapeStyle).not.toHaveBeenCalled();

		gallery.update(state({ canShape: true }));
		expect(gallery.el.hidden).toBeFalsy();
		expect(tiles().some((tile) => tile.disabled)).toBeFalsy();
	});
});

describe('text 3d section', () => {
	function build() {
		const setTextStyle = vi.fn();
		const section = createText3DSection(document, createTranslator(), sectionFactory(), {
			setTextStyle,
		} as unknown as InspectorHandlers);
		return { section, setTextStyle };
	}

	it('seeds a 6pt extrusion when the toggle is switched on', () => {
		const { section, setTextStyle } = build();
		section.update(state({ canText: true, textStyle: {} as TextStyle }));

		const toggle = section.el.querySelector<HTMLInputElement>('pptx-ui-checkbox')!;
		toggle.checked = true;
		toggle.dispatchEvent(new Event('change'));

		expect(setTextStyle).toHaveBeenCalledWith({ text3d: { extrusionHeight: 76200 } });
	});

	it('clears the whole 3d style when the toggle is switched off', () => {
		const { section, setTextStyle } = build();
		section.update(
			state({ canText: true, textStyle: { text3d: { extrusionHeight: 76200 } } as TextStyle }),
		);

		const toggle = section.el.querySelector<HTMLInputElement>('pptx-ui-checkbox')!;
		toggle.checked = false;
		toggle.dispatchEvent(new Event('change'));

		expect(setTextStyle).toHaveBeenCalledWith({ text3d: undefined });
	});

	it('shows the bevel/material options only once there is extrusion', () => {
		const { section } = build();
		section.update(state({ canText: true, textStyle: {} as TextStyle }));
		const options = section.el.querySelector<HTMLElement>('.pptxv-text3d-options')!;
		expect(options.hidden).toBeTruthy();

		section.update(
			state({ canText: true, textStyle: { text3d: { extrusionHeight: 76200 } } as TextStyle }),
		);
		expect(options.hidden).toBeFalsy();
	});

	it('commits a bevel width in EMU', () => {
		const { section, setTextStyle } = build();
		section.update(
			state({ canText: true, textStyle: { text3d: { extrusionHeight: 76200 } } as TextStyle }),
		);

		const widths = section.el.querySelectorAll<HTMLInputElement>('.pptxv-text3d-bevel input');
		const width = Array.from(widths).find((input) => input.type === 'number')!;
		width.value = '3';
		width.dispatchEvent(new Event('change'));

		expect(setTextStyle).toHaveBeenCalledWith(
			expect.objectContaining({ text3d: expect.objectContaining({ bevelTopWidth: 38100 }) }),
		);
	});
});

describe('media trim timeline', () => {
	it('projects the trim window and playhead onto the bar', () => {
		const timeline = createMediaTrimTimeline(document, {
			onTrimChange: vi.fn(),
			onSeek: vi.fn(),
		});
		// `trimEndMs` is p14:trim/@end's distance from the tail (G19): 2000 off
		// a 10s clip ends at 8s, so the window runs 2s..8s.
		timeline.update({
			duration: 10,
			trimStartMs: 2000,
			trimEndMs: 2000,
			currentTime: 5,
			bookmarks: [{ id: 'b1', time: 5, label: 'Mid' }],
			canEdit: true,
		});

		const region = timeline.el.querySelector<HTMLElement>('.pptxv-media-timeline-region')!;
		expect(region.style.left).toBe('20%');
		expect(region.style.width).toBe('60%');
		const labels = timeline.el.querySelectorAll<HTMLElement>('.pptxv-media-timeline-time');
		expect(labels[0]?.textContent).toBe('0:02.0');
		expect(labels[1]?.textContent).toBe('0:08.0');
		expect(
			timeline.el.querySelector<HTMLElement>('.pptxv-media-timeline-playhead')!.style.left,
		).toBe('50%');
		expect(timeline.el.querySelectorAll('.pptxv-media-timeline-mark')).toHaveLength(1);
	});

	it('seeks to a bookmark when its marker is clicked', () => {
		const onSeek = vi.fn();
		const timeline = createMediaTrimTimeline(document, { onTrimChange: vi.fn(), onSeek });
		timeline.update({
			duration: 10,
			trimStartMs: 0,
			trimEndMs: 0,
			currentTime: 0,
			bookmarks: [{ id: 'b1', time: 4, label: 'Cue' }],
			canEdit: true,
		});

		timeline.el.querySelector<HTMLButtonElement>('.pptxv-media-timeline-mark')!.click();
		expect(onSeek).toHaveBeenCalledWith(4);
	});

	it('hides the drag handles when the deck is read-only', () => {
		const timeline = createMediaTrimTimeline(document, {
			onTrimChange: vi.fn(),
			onSeek: vi.fn(),
		});
		timeline.update({
			duration: 4,
			trimStartMs: 0,
			trimEndMs: 0,
			currentTime: 0,
			bookmarks: [],
			canEdit: false,
		});
		expect(
			timeline.el.querySelector<HTMLElement>('.pptxv-media-timeline-handle.is-start')!.hidden,
		).toBeTruthy();
	});
});

describe('image alt text', () => {
	it('commits the accessibility description on change', () => {
		const setAltText = vi.fn();
		const handlers = {
			setAltText,
			replaceImage: vi.fn(),
			resetImage: vi.fn(),
			setImageBrightness: vi.fn(),
			setImageContrast: vi.fn(),
			setImageSaturation: vi.fn(),
			setImageCrop: vi.fn(),
			setImageEffects: vi.fn(),
		} as unknown as InspectorHandlers;
		const section = createImageSection(document, createTranslator(), sectionFactory(), handlers);
		section.update(state({ isImage: true, altText: 'Old' }));

		const alt = section.el.querySelector<HTMLTextAreaElement>('.pptxv-image-alt-input')!;
		expect(alt.value).toBe('Old');
		alt.value = 'A cat on a mat';
		alt.dispatchEvent(new Event('change'));

		expect(setAltText).toHaveBeenCalledWith('A cat on a mat');
	});
});

describe('media section trim timeline wiring', () => {
	it('renders a timeline inside the media section', async () => {
		const { createMediaSection } = await import('./media-section');
		const section = createMediaSection(document, createTranslator(), sectionFactory(), {
			setMediaProperties: vi.fn(),
		} as unknown as InspectorHandlers);
		section.update(
			state({
				isMedia: true,
				media: { type: 'media', id: 'm1', trimStartMs: 0, trimEndMs: 0 } as MediaPptxElement,
			}),
		);
		expect(section.el.querySelector('.pptxv-media-timeline')).not.toBeNull();
	});

	/**
	 * `trimEndMs` is `p14:trim/@end`'s distance from the clip's TAIL
	 * (COM-verified). The "Trim End" number field used to bind that distance
	 * directly, so typing "the last 5s" of a 20s clip meant computing
	 * 20000-5000 by hand; it now shows/accepts the absolute end position, like
	 * React's `MediaInspector` and Vue's `MediaPropertiesPanel.vue`.
	 */
	it('shows 15000 (15s) for a 20s clip with trimEndMs=5000, and stores 5000 for a typed 15000', async () => {
		const { createMediaSection } = await import('./media-section');
		const setMediaProperties = vi.fn();
		const section = createMediaSection(document, createTranslator(), sectionFactory(), {
			setMediaProperties,
		} as unknown as InspectorHandlers);
		section.update(
			state({
				isMedia: true,
				media: {
					type: 'media',
					id: 'm1',
					trimStartMs: 0,
					trimEndMs: 5000,
					metadata: { duration: 20 },
				} as MediaPptxElement,
			}),
		);

		const trimEndInput = section.el.querySelector<HTMLInputElement>(
			'input[aria-label="Trim End"]',
		)!;
		expect(trimEndInput.value).toBe('15000');

		// A different absolute end than the one just displayed (15000), so the
		// field's own "unchanged commit" guard does not swallow it.
		trimEndInput.value = '12000';
		trimEndInput.dispatchEvent(new Event('change'));
		expect(setMediaProperties).toHaveBeenCalledWith({ trimEndMs: 8000 });
	});
});
