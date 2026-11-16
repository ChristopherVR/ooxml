import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createAccessibilitySection } from './accessibility-section';
import type { InspectorHandlers, InspectorState } from './types';

function sectionFactory() {
	return (): HTMLElement => document.createElement('div');
}

function shape(extra: Record<string, unknown> = {}): PptxElement {
	return { type: 'shape', id: 's1', x: 0, y: 0, width: 10, height: 10, ...extra } as PptxElement;
}

function state(overrides: Partial<InspectorState> = {}): InspectorState {
	return {
		hasSelection: true,
		showAccessibilitySection: true,
		selectedElement: shape(),
		...overrides,
	} as InspectorState;
}

function handlers(): InspectorHandlers {
	return { patchSelectedElement: vi.fn() } as unknown as InspectorHandlers;
}

function create(h = handlers()) {
	return createAccessibilitySection(document, createTranslator(), sectionFactory(), h);
}

describe('accessibility section (alt text and title)', () => {
	it('is hidden when nothing is selected', () => {
		const section = create();
		section.update(state({ hasSelection: false }));
		expect(section.el.hidden).toBeTruthy();
	});

	it('is hidden for an element kind the shared descriptor excludes (e.g. a picture or group)', () => {
		const section = create();
		section.update(state({ showAccessibilitySection: false }));
		expect(section.el.hidden).toBeTruthy();
	});

	it('shows the current altText and title for a shape', () => {
		const section = create();
		section.update(
			state({ selectedElement: shape({ altText: 'A red rectangle', title: 'Callout' }) }),
		);
		expect(section.el.hidden).toBeFalsy();
		const textarea = section.el.querySelector('textarea') as HTMLTextAreaElement;
		const input = section.el.querySelector('input[type="text"]') as HTMLInputElement;
		expect(textarea.value).toBe('A red rectangle');
		expect(input.value).toBe('Callout');
	});

	it('commits altText and title as they are typed', () => {
		const h = handlers();
		const section = create(h);
		section.update(state());
		const textarea = section.el.querySelector('textarea') as HTMLTextAreaElement;
		textarea.value = 'Updated description';
		textarea.dispatchEvent(new Event('input'));
		expect(h.patchSelectedElement).toHaveBeenCalledWith({ altText: 'Updated description' });
		const input = section.el.querySelector('input[type="text"]') as HTMLInputElement;
		input.value = 'Updated title';
		input.dispatchEvent(new Event('input'));
		expect(h.patchSelectedElement).toHaveBeenCalledWith({ title: 'Updated title' });
	});
});
