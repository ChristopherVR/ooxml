/**
 * The Properties tab as a selection-aware Format pane: the shared section list
 * for the selection, open state kept per category in the viewer's controller,
 * and launcher requests that expand and focus a section.
 */
import { mount } from '@vue/test-utils';
import { createFormatPaneController } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { nextTick } from 'vue';

import { FormatPaneKey } from '../../composables/format-pane-context';
import InspectorPane from './InspectorPane.vue';

type SectionHost = HTMLElement & { open: boolean; label: string; mixed: boolean };

const shape = (id: string, fillColor = '#ff0000'): PptxElement =>
	({
		id,
		type: 'shape',
		shapeType: 'rect',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeStyle: { fillColor },
		textStyle: {},
	}) as PptxElement;

let wrappers: { unmount(): void }[] = [];
afterEach(() => {
	for (const w of wrappers) {
		w.unmount();
	}
	wrappers = [];
});

function render(pane: FormatPaneController, element: PptxElement, selection?: PptxElement[]) {
	const wrapper = mount(InspectorPane, {
		props: { element, ...(selection ? { selection } : {}) },
		attachTo: document.body,
		global: { provide: { [FormatPaneKey as symbol]: pane } },
	});
	wrappers.push(wrapper);
	return wrapper;
}

const section = (root: Element, id: string) =>
	root.querySelector<SectionHost>(`[data-pptx-format-section="${id}"]`);
const sectionIds = (root: Element) =>
	[...root.querySelectorAll<HTMLElement>('[data-pptx-format-section]')].map(
		(node) => node.dataset.pptxFormatSection,
	);

describe('Vue Format pane sections', () => {
	it('orders a shape as Fill, Outline open, Effects closed, with panels inside', () => {
		const root = render(createFormatPaneController(), shape('a')).element;
		expect(sectionIds(root).slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(section(root, 'fill')!.open).toBe(true);
		expect(section(root, 'outline')!.open).toBe(true);
		expect(section(root, 'effects')!.open).toBe(false);
		expect(section(root, 'sizePosition')!.open).toBe(false);
		expect(section(root, 'fill')!.label).toBe('Fill');
		expect(
			section(root, 'fill')!.querySelector('[data-pptx-inspector-section="fill-stroke"]'),
		).not.toBeNull();
		expect(
			section(root, 'sizePosition')!.querySelector('[data-pptx-inspector-section="transform"]'),
		).not.toBeNull();
	});

	it('remembers a toggled section for the next element of the same category', async () => {
		const pane = createFormatPaneController();
		const wrapper = render(pane, shape('a'));
		section(wrapper.element, 'fill')!.shadowRoot!.querySelector('summary')!.click();
		await nextTick();
		expect(section(wrapper.element, 'fill')!.open).toBe(false);
		await wrapper.setProps({ element: shape('b') });
		expect(section(wrapper.element, 'fill')!.open).toBe(false);
		expect(pane.isOpen('shape', { id: 'fill', labelKey: '', defaultOpen: true })).toBe(false);
	});

	it('expands and focuses a requested section', async () => {
		vi.useFakeTimers({ toFake: ['requestAnimationFrame'] });
		try {
			const pane = createFormatPaneController();
			const wrapper = render(pane, shape('a'));
			pane.requestFormatPane('sizePosition', { focus: true });
			await nextTick();
			await nextTick();
			const target = section(wrapper.element, 'sizePosition')!;
			expect(target.open).toBe(true);
			vi.advanceTimersToNextFrame();
			expect(target.shadowRoot!.activeElement).toBe(target.shadowRoot!.querySelector('summary'));
		} finally {
			vi.useRealTimers();
		}
	});

	it('drops Size and position and tags differing values for a multi-selection', () => {
		const a = shape('a', '#ff0000');
		const root = render(createFormatPaneController(), a, [a, shape('b', '#00ff00')]).element;
		expect(sectionIds(root)).not.toContain('sizePosition');
		expect(section(root, 'fill')!.mixed).toBeTruthy();
		expect(section(root, 'outline')!.mixed).toBeFalsy();
	});
});
