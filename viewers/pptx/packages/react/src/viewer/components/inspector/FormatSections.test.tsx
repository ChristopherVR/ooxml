// @vitest-environment happy-dom
/**
 * The Properties tab as a selection-aware Format pane: the shared section list
 * for the selection, open state kept per category in the viewer's controller,
 * and launcher requests that expand and focus a section.
 */
import { createFormatPaneController } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { ElementInspectorBody } from './ElementInspectorBody';
import { FormatPaneContext } from './format-pane-context';

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

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function render(
	pane: FormatPaneController,
	selected: PptxElement,
	selection: PptxElement[] = [selected],
): void {
	act(() => {
		root.render(
			<FormatPaneContext.Provider value={pane}>
				<ElementInspectorBody
					selectedElement={selected}
					selectedElements={selection}
					canEdit
					slides={[]}
					customShows={[]}
					onUpdateElement={vi.fn()}
					onUpdateElementStyle={vi.fn()}
					onUpdateTextStyle={vi.fn()}
					onMoveLayer={vi.fn()}
				/>
			</FormatPaneContext.Provider>,
		);
	});
}

const section = (id: string) =>
	container.querySelector<SectionHost>(`[data-pptx-format-section="${id}"]`);
const sectionIds = () =>
	[...container.querySelectorAll<HTMLElement>('[data-pptx-format-section]')].map(
		(node) => node.dataset.pptxFormatSection,
	);

describe('React Format pane sections', () => {
	it('orders a shape as Fill, Outline open, Effects closed, with panels inside', () => {
		render(createFormatPaneController(), shape('a'));
		expect(sectionIds().slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(section('fill')!.open).toBe(true);
		expect(section('outline')!.open).toBe(true);
		expect(section('effects')!.open).toBe(false);
		expect(section('sizePosition')!.open).toBe(false);
		expect(section('fill')!.querySelector('[data-pptx-fill-stroke]')).not.toBeNull();
		expect(
			section('sizePosition')!.querySelector('[data-pptx-inspector-section="transform"]'),
		).not.toBeNull();
	});

	it('remembers a toggled section for the next element of the same category', () => {
		const pane = createFormatPaneController();
		render(pane, shape('a'));
		act(() => {
			section('fill')!.shadowRoot!.querySelector('summary')!.click();
		});
		expect(section('fill')!.open).toBe(false);
		render(pane, shape('b'));
		expect(section('fill')!.open).toBe(false);
		expect(pane.isOpen('shape', { id: 'fill', labelKey: '', defaultOpen: true })).toBe(false);
	});

	it('expands and focuses a requested section', () => {
		vi.useFakeTimers({ toFake: ['requestAnimationFrame'] });
		try {
			const pane = createFormatPaneController();
			render(pane, shape('a'));
			act(() => pane.requestFormatPane('sizePosition', { focus: true }));
			expect(section('sizePosition')!.open).toBe(true);
			act(() => {
				vi.advanceTimersToNextFrame();
			});
			const summary = section('sizePosition')!.shadowRoot!.querySelector('summary');
			expect(section('sizePosition')!.shadowRoot!.activeElement).toBe(summary);
		} finally {
			vi.useRealTimers();
		}
	});

	it('drops Size and position and tags differing values for a multi-selection', () => {
		const a = shape('a', '#ff0000');
		render(createFormatPaneController(), a, [a, shape('b', '#00ff00')]);
		expect(sectionIds()).not.toContain('sizePosition');
		expect(section('fill')!.mixed).toBe(true);
		expect(section('outline')!.mixed).toBe(false);
	});
});
