// @vitest-environment happy-dom
/**
 * The React generic field renderer draws the shared inspector descriptors
 * (Effects, Line, Size & Position, Text Box): one toggle per toggleable
 * section, its fields only while on, and every change goes out as the
 * descriptor's patch.
 */
import {
	ELEMENT_ALT_TEXT_SECTIONS,
	ELEMENT_GEOMETRY_SECTIONS,
	PICTURE_ADJUSTMENT_SECTIONS,
	SHAPE_EFFECT_SECTIONS,
	SHAPE_LINE_SECTIONS,
	TEXT_BOX_SECTIONS,
} from 'ooxml-ui/pptx';
import type {
	MediaPptxElement,
	PptxElement,
	PptxImageLikeElement,
	ShapeStyle,
	TextStyle,
} from 'pptx-viewer-core';
import React, { act } from 'react';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createMarkedRoot, marked } from '../../i18n/marking-translator.test-support';
import { ImagePropertiesPanel } from './ImagePropertiesPanel';
import { InspectorFields } from './InspectorFields';
import { MediaPlaybackBookmarks } from './MediaPlaybackBookmarks';
import { TextAdvancedSections } from './TextAdvancedSections';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createMarkedRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function render(style: ShapeStyle | undefined, onPatch: (p: Partial<ShapeStyle>) => void) {
	act(() => {
		root.render(
			<div className='grid grid-cols-2'>
				<InspectorFields sections={SHAPE_EFFECT_SECTIONS} target={style} onPatch={onPatch} />
			</div>,
		);
	});
}

const byTestId = <E extends Element = HTMLInputElement>(id: string) =>
	container.querySelector<E & HTMLInputElement>(`[data-testid="${id}"]`);

describe('inspectorFields (effects)', () => {
	it('draws a toggle per effect section and no fields while every effect is off', () => {
		render({}, vi.fn());
		for (const id of [
			'fx-outer-shadow-toggle',
			'fx-inner-shadow-toggle',
			'fx-glow-toggle',
			'fx-soft-edge-toggle',
			'fx-reflection-toggle',
			'fx-blur-toggle',
			'fx-bevel-toggle',
		]) {
			expect(byTestId(id)?.tagName.toLowerCase()).toBe('pptx-ui-checkbox');
		}
		expect(byTestId('fx-outer-shadow-blur')).toBeNull();
	});

	it('emits the enable patch when a toggle is ticked', () => {
		const onPatch = vi.fn();
		render({}, onPatch);
		const toggle = byTestId('fx-glow-toggle')!;
		act(() => {
			toggle.checked = true;
			toggle.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({
			glowColor: '#ffff00',
			glowRadius: 6,
			glowOpacity: 0.75,
		});
	});

	it("shows an enabled effect's fields and writes a clamped number", () => {
		const onPatch = vi.fn();
		render({ shadowColor: '#112233', shadowBlur: 3, shadowAngle: 0, shadowDistance: 4 }, onPatch);
		const blur = byTestId('fx-outer-shadow-blur')!;
		expect(blur.value).toBe('3');
		expect(blur.getAttribute('aria-label')).toBeTruthy();
		act(() => {
			const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
			setter.call(blur, '500');
			blur.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith(expect.objectContaining({ shadowBlur: 96 }));
		expect(byTestId('fx-outer-shadow-rotate-with-shape')).not.toBeNull();
	});
});

describe('inspectorFields (line options)', () => {
	function renderLine(style: ShapeStyle | undefined, onPatch: (p: Partial<ShapeStyle>) => void) {
		act(() => {
			root.render(
				<div className='grid grid-cols-2'>
					<InspectorFields sections={SHAPE_LINE_SECTIONS} target={style} onPatch={onPatch} />
				</div>,
			);
		});
	}

	it('draws the untitled line fields with their current values', () => {
		renderLine({ strokeWidth: 3, strokeOpacity: 0.4, lineCap: 'sq' }, vi.fn());
		expect(byTestId('line-width')!.value).toBe('3');
		expect(byTestId('line-opacity')!.value).toBe('40');
		expect(byTestId('line-cap')!.tagName.toLowerCase()).toBe('pptx-ui-select');
		expect(container.querySelector('pptx-ui-checkbox')).toBeNull();
	});

	it('writes a dash choice as a shape-style patch', () => {
		const onPatch = vi.fn();
		renderLine({}, onPatch);
		const dash = byTestId('line-dash')!;
		act(() => {
			dash.value = 'dash';
			dash.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({ strokeDash: 'dash' });
	});
});

describe('inspectorFields (size and position)', () => {
	it('shows the frame in whole px and writes a geometry patch, disabled when read-only', () => {
		const onPatch = vi.fn();
		const frame = { x: 53.4, y: 10, width: 200, height: 100, rotation: 0 };
		act(() => {
			root.render(
				<InspectorFields sections={ELEMENT_GEOMETRY_SECTIONS} target={frame} onPatch={onPatch} />,
			);
		});
		const x = byTestId('geometry-x')!;
		expect(x.value).toBe('53');
		expect(x.getAttribute('aria-label')).toBe(marked('pptx.inspector.x'));
		act(() => {
			const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
			setter.call(byTestId('geometry-width')!, '0');
			byTestId('geometry-width')!.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({ width: 1 });

		act(() => {
			root.render(
				<InspectorFields
					sections={ELEMENT_GEOMETRY_SECTIONS}
					target={frame}
					onPatch={onPatch}
					disabled
				/>,
			);
		});
		expect(byTestId('geometry-rotation')!.disabled).toBeTruthy();
	});
});

describe('inspectorFields (text box options)', () => {
	it('draws the text box options and writes body-level text-style patches', () => {
		const onPatch = vi.fn<(p: Partial<TextStyle>) => void>();
		act(() => {
			root.render(
				<InspectorFields
					sections={TEXT_BOX_SECTIONS}
					target={{ vAlign: 'bottom' } as TextStyle}
					onPatch={onPatch}
				/>,
			);
		});
		expect(byTestId('text-box-valign')!.value).toBe('bottom');
		const wrap = byTestId('text-box-wrap')!;
		act(() => {
			wrap.checked = false;
			wrap.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({ textWrap: 'none' });
		const autofit = byTestId('text-box-autofit')!;
		act(() => {
			autofit.value = 'normal';
			autofit.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenLastCalledWith({ autoFitMode: 'normal', autoFit: true });
	});
});

describe('inspectorFields (alt text and picture)', () => {
	const picture = {
		type: 'image',
		id: 'p1',
		x: 0,
		y: 0,
		width: 10,
		height: 10,
		imageEffects: { alphaModFix: 60 },
	} as unknown as PptxElement;

	function renderPicture(onPatch: (p: Partial<PptxElement>) => void, disabled = false) {
		act(() => {
			root.render(
				<div className='grid grid-cols-2'>
					<InspectorFields
						sections={ELEMENT_ALT_TEXT_SECTIONS}
						target={picture}
						onPatch={onPatch}
						disabled={disabled}
					/>
					<InspectorFields
						sections={PICTURE_ADJUSTMENT_SECTIONS}
						target={picture}
						onPatch={onPatch}
						disabled={disabled}
					/>
				</div>,
			);
		});
	}

	it('draws the description as a text area and writes what is typed', () => {
		const onPatch = vi.fn();
		renderPicture(onPatch);
		const alt = byTestId<HTMLTextAreaElement>('element-alt-text')!;
		expect(alt.tagName.toLowerCase()).toBe('textarea');
		expect(byTestId('element-alt-title')).toBeNull();
		act(() => {
			const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!;
			setter.call(alt, 'A sunset');
			alt.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(onPatch).toHaveBeenCalledWith({ altText: 'A sunset' });
	});

	it('shows transparency inverted from the stored opacity and disables with the pane', () => {
		renderPicture(vi.fn(), true);
		const transparency = byTestId('picture-transparency')!;
		expect(transparency.value).toBe('40');
		expect(transparency.disabled).toBeTruthy();
		expect(byTestId<HTMLTextAreaElement>('element-alt-text')!.disabled).toBeTruthy();
	});
});

describe('textAdvancedSections spacing fields', () => {
	it('draws the shared paragraph spacing in points and writes px', () => {
		const onUpdateTextStyle = vi.fn();
		const element = {
			type: 'text',
			id: 't1',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			textStyle: { paragraphMarginLeft: 48 },
		} as unknown as PptxElement;
		act(() => {
			root.render(
				<TextAdvancedSections
					selectedElement={element}
					canEdit
					onUpdateTextStyle={onUpdateTextStyle}
					part='paragraph'
				/>,
			);
		});
		const margin = byTestId('paragraph-marginLeft')!;
		expect(margin.value).toBe('36');
		expect(byTestId('font-characterSpacing')).toBeNull();
		act(() => {
			const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
			setter.call(margin, '18');
			margin.dispatchEvent(new Event('input', { bubbles: true }));
		});
		expect(onUpdateTextStyle).toHaveBeenCalledWith({ paragraphMarginLeft: 24 });
	});
});

describe('imagePropertiesPanel colour effects', () => {
	it('draws the shared colour effects and switches a colour wash on', () => {
		const onUpdateElement = vi.fn();
		const picture = {
			type: 'image',
			id: 'p1',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			imageEffects: { brightness: 5 },
		} as unknown as PptxImageLikeElement;
		act(() => {
			root.render(
				<ImagePropertiesPanel
					selectedElement={picture}
					canEdit
					onUpdateElement={onUpdateElement}
				/>,
			);
		});
		expect(byTestId('picture-washColor')).toBeNull();
		const wash = byTestId('picture-colorWash-toggle')!;
		act(() => {
			wash.checked = true;
			wash.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onUpdateElement).toHaveBeenCalledWith({
			imageEffects: { brightness: 5, colorWash: { color: '#0066cc', opacity: 40 } },
		});
		expect(
			container.querySelectorAll('[data-testid="picture-duotone-preset"]').length,
		).toBeGreaterThan(0);
	});
});

describe('mediaPlaybackBookmarks playback fields', () => {
	it('draws the shared playback options and writes a media patch', () => {
		const onUpdateElement = vi.fn();
		const video = {
			type: 'media',
			id: 'm1',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			mediaType: 'video',
		} as unknown as MediaPptxElement;
		act(() => {
			root.render(
				<MediaPlaybackBookmarks
					element={video}
					canEdit
					currentTime={0}
					bookmarks={[]}
					onUpdateElement={onUpdateElement}
					onSeekTo={vi.fn()}
				/>,
			);
		});
		expect(byTestId('media-playAcrossSlides')).toBeNull();
		const loop = byTestId('media-loop')!;
		act(() => {
			loop.checked = true;
			loop.dispatchEvent(new Event('change', { bubbles: true }));
		});
		expect(onUpdateElement).toHaveBeenCalledWith({ loop: true });
	});
});
