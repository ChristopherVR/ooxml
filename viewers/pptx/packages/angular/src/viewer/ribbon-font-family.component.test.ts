/**
 * Choosing a font family with the whole text box selected (no inline
 * selection) must reach every run, the way the font size already did.
 * Patching only the body left runs that name their own font unchanged.
 */
import { Injector, runInInjectionContext } from '@angular/core';
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EditorStateService } from './editor-state.service';
import { RibbonFontControlsComponent } from './ribbon-font-controls.component';

const sameFace = { fontFamily: 'Old Sans', eastAsiaFont: 'Old Sans' };

/** A run that inherits its Latin face from the box but names the old face in `a:ea`. */
const inheritsLatin = { eastAsiaFont: 'Old Sans', bold: true };

function textBox(runStyle: TextSegment['style'] = sameFace): PptxElement {
	return {
		type: 'text',
		id: 'box',
		x: 0,
		y: 0,
		width: 100,
		height: 40,
		text: 'Sample',
		textStyle: { ...sameFace },
		textSegments: [{ text: 'Sample', style: { ...runStyle } }],
	} as PptxElement;
}

function pickFontFamily(element: PptxElement): ReturnType<typeof vi.fn> {
	const updateElement = vi.fn();
	const injector = Injector.create({
		providers: [{ provide: EditorStateService, useValue: { updateElement } }],
	});
	const component = runInInjectionContext(injector, () => new RibbonFontControlsComponent());
	Object.defineProperty(component, 'selectedElement', { value: () => element });

	(component as unknown as { pickerRequest(event: Event): void }).pickerRequest(
		new CustomEvent('request', { detail: { id: 'home.font.fontFamily', value: 'New Serif' } }),
	);
	return updateElement;
}

describe('ribbonFontControlsComponent font family on a whole box', () => {
	it('applies the family to every run, east Asian face included', () => {
		const updateElement = pickFontFamily(textBox());

		expect(updateElement).toHaveBeenCalledOnce();
		const patch = updateElement.mock.calls[0]?.[2] as { textSegments: TextSegment[] };
		expect(patch.textSegments[0]?.style).toMatchObject({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
		});
	});

	it('carries an east Asian face that matched the inherited Latin face', () => {
		const updateElement = pickFontFamily(textBox(inheritsLatin));

		const patch = updateElement.mock.calls[0]?.[2] as { textSegments: TextSegment[] };
		expect(patch.textSegments[0]?.style).toStrictEqual({
			fontFamily: 'New Serif',
			eastAsiaFont: 'New Serif',
			bold: true,
		});
	});
});
