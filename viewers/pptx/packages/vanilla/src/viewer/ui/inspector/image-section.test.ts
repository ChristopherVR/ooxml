import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createImageSection } from './image-section';
import type { InspectorHandlers, InspectorState } from './types';

function sectionFactory() {
	return (): HTMLElement => document.createElement('div');
}

function picture(extra: Record<string, unknown> = {}): PptxElement {
	return { type: 'image', id: 'p1', x: 0, y: 0, width: 10, height: 10, ...extra } as PptxElement;
}

function state(overrides: Partial<InspectorState> = {}): InspectorState {
	return {
		hasSelection: true,
		isImage: true,
		croppable: true,
		selectedElement: picture(),
		...overrides,
	} as InspectorState;
}

function handlers(): InspectorHandlers {
	return {
		patchSelectedElement: vi.fn(),
		replaceImage: vi.fn(),
		resetImage: vi.fn(),
		setImageEffects: vi.fn(),
		pushRecentColor: vi.fn(),
	} as unknown as InspectorHandlers;
}

const cropInputs = (el: HTMLElement) =>
	el.querySelectorAll<HTMLInputElement>('[data-testid^="picture-crop"]');

/**
 * G7 (OpenXML parity audit, D3): `a:picLocks/@noCrop` was parsed and
 * round-tripped but never enforced - the four crop fields stayed enabled
 * regardless of the lock, gated only on `state.isImage`.
 */
describe('image section crop fields with a:picLocks/@noCrop', () => {
	it('disables all four crop fields when croppable is false', () => {
		const section = createImageSection(document, createTranslator(), sectionFactory(), handlers());
		section.update(state({ croppable: false }));
		expect(cropInputs(section.el)).toHaveLength(4);
		for (const input of cropInputs(section.el)) {
			expect(input.disabled).toBeTruthy();
		}
	});

	it('leaves the crop fields enabled on an editable, unlocked picture', () => {
		const section = createImageSection(document, createTranslator(), sectionFactory(), handlers());
		section.update(state({ croppable: true }));
		for (const input of cropInputs(section.el)) {
			expect(input.disabled).toBeFalsy();
		}
	});
});

describe('image section shared picture fields', () => {
	it('shows transparency inverted and commits a crop inset as a fraction', () => {
		const h = handlers();
		const section = createImageSection(document, createTranslator(), sectionFactory(), h);
		section.update(state({ selectedElement: picture({ imageEffects: { alphaModFix: 80 } }) }));
		const transparency = section.el.querySelector<HTMLInputElement>(
			'[data-testid="picture-transparency"]',
		)!;
		expect(transparency.value).toBe('20');
		const left = section.el.querySelector<HTMLInputElement>('[data-testid="picture-cropLeft"]')!;
		left.value = '25';
		left.dispatchEvent(new Event('change'));
		expect(h.patchSelectedElement).toHaveBeenCalledWith({ cropLeft: 0.25 });
	});

	it('commits the alt text as it is typed', () => {
		const h = handlers();
		const section = createImageSection(document, createTranslator(), sectionFactory(), h);
		section.update(state({ selectedElement: picture({ altText: 'Old' }) }));
		const alt = section.el.querySelector<HTMLTextAreaElement>('[data-testid="element-alt-text"]')!;
		expect(alt.value).toBe('Old');
		alt.value = 'A cat on a mat';
		alt.dispatchEvent(new Event('input'));
		expect(h.patchSelectedElement).toHaveBeenCalledWith({ altText: 'A cat on a mat' });
	});
});
