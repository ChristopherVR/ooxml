import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import { ChromeUiState } from '../../state/chrome-ui.svelte';
import InspectorPanel from './InspectorPanel.svelte';

/**
 * The Properties tab as a selection-aware Format pane: the shared section
 * list for the selection, open state kept per category in the viewer's
 * controller (`ChromeUiState.formatPane`), and launcher requests that open
 * the tab and expand and focus a section.
 */

type SectionHost = HTMLElement & { open: boolean; mixed: boolean };

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const shape = (id: string, fillColor = '#ff0000'): PptxElement =>
	({
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		shapeType: 'rect',
		shapeStyle: { fillColor },
		textStyle: {},
	}) as PptxElement;

function setup(elements: PptxElement[]): {
	editor: EditorState;
	ui: ChromeUiState;
	target: HTMLElement;
} {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements }]);
	const ui = new ChromeUiState(true);
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(InspectorPanel, { target, props: { editor, ui } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return { editor, ui, target };
}

function select(editor: EditorState, ...ids: string[]): void {
	editor.selection.setAll(ids);
	flushSync();
}

const section = (target: HTMLElement, id: string) =>
	target.querySelector<SectionHost>(`[data-pptx-format-section="${id}"]`);
const sectionIds = (target: HTMLElement) =>
	[...target.querySelectorAll<HTMLElement>('[data-pptx-format-section]')].map(
		(node) => node.dataset.pptxFormatSection,
	);

describe('Svelte Format pane sections', () => {
	it('orders a shape as Fill, Outline open, Effects closed, with panels inside', () => {
		const { editor, target } = setup([shape('a')]);
		select(editor, 'a');
		expect(sectionIds(target).slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(section(target, 'fill')!.open).toBe(true);
		expect(section(target, 'outline')!.open).toBe(true);
		expect(section(target, 'effects')!.open).toBe(false);
		expect(section(target, 'sizePosition')!.open).toBe(false);
		expect(
			section(target, 'fill')!.querySelector('[data-pptx-inspector-section="fill-stroke"]'),
		).not.toBeNull();
		expect(
			section(target, 'sizePosition')!.querySelector('[data-pptx-inspector-section="transform"]'),
		).not.toBeNull();
	});

	it('remembers a toggled section for the next element of the same category', () => {
		const { editor, ui, target } = setup([shape('a'), shape('b')]);
		select(editor, 'a');
		section(target, 'fill')!.shadowRoot!.querySelector('summary')!.click();
		flushSync();
		expect(section(target, 'fill')!.open).toBe(false);
		select(editor, 'b');
		expect(section(target, 'fill')!.open).toBe(false);
		expect(ui.formatPane.isOpen('shape', { id: 'fill', labelKey: '', defaultOpen: true })).toBe(
			false,
		);
	});

	it('switches to Properties, expands and focuses a requested section', () => {
		vi.useFakeTimers({ toFake: ['requestAnimationFrame'] });
		try {
			const { editor, ui, target } = setup([shape('a')]);
			select(editor, 'a');
			ui.setInspectorTab('comments');
			flushSync();
			ui.formatPane.requestFormatPane('sizePosition', { focus: true });
			flushSync();
			expect(ui.inspectorTab).toBe('properties');
			expect(section(target, 'sizePosition')!.open).toBe(true);
			vi.advanceTimersToNextFrame();
			const host = section(target, 'sizePosition')!;
			expect(host.shadowRoot!.activeElement).toBe(host.shadowRoot!.querySelector('summary'));
		} finally {
			vi.useRealTimers();
		}
	});

	it('drops Size and position and tags differing values for a multi-selection', () => {
		const { editor, target } = setup([shape('a', '#ff0000'), shape('b', '#00ff00')]);
		select(editor, 'a', 'b');
		expect(sectionIds(target)).not.toContain('sizePosition');
		expect(section(target, 'fill')!.mixed).toBe(true);
		expect(section(target, 'outline')!.mixed).toBe(false);
	});
});
