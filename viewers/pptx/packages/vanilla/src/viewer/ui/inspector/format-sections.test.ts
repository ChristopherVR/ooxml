/**
 * The vanilla Properties tab as a selection-aware Format pane: the shared
 * section list for the selection, open state kept per category in the
 * viewer's controller, and launcher requests that expand and focus a section.
 */
import { createFormatPaneController, registerPptxWebControls } from 'ooxml-ui/pptx';
import type { FormatPaneController } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { buildInspectorState } from '../../editor/inspector-state-builder';
import { createInspector } from './inspector';
import type { InspectorDeckState, InspectorHandlers } from './types';

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
	}) as unknown as PptxElement;

beforeAll(() => registerPptxWebControls());
afterEach(() => document.body.replaceChildren());

function mount(pane: FormatPaneController) {
	const handlers = new Proxy({ formatPane: pane } as unknown as InspectorHandlers, {
		get: (target, key, receiver) =>
			Reflect.get(target, key, receiver) ?? (key === 'then' ? undefined : vi.fn()),
	});
	const inspector = createInspector(document, (key) => key, handlers);
	document.body.appendChild(inspector.el);
	const select = (primary: PptxElement, selection: PptxElement[] = [primary]): void => {
		inspector.update(buildInspectorState(primary));
		inspector.updateDeck({
			slideCount: 1,
			currentSlide: 0,
			canvasSize: { width: 960, height: 540 },
			slideSize: undefined,
			hasDeckElements: true,
			elements: selection,
			selectedIds: selection.map((el) => el.id),
			selectedElementId: primary.id,
			selectedElement: primary,
			comments: [],
			commentMentionAuthors: [],
			customShows: [],
			docTitle: undefined,
			docAuthor: undefined,
			editable: true,
			presentationProperties: {},
			themeOptions: [],
			activeSlide: undefined,
			editTemplateMode: false,
			slideMasters: [],
			colorScheme: undefined,
			fontScheme: undefined,
			themeName: undefined,
			tagCollections: [],
			notesCanvasSize: undefined,
			notesPlaceholderCount: undefined,
			handoutPlaceholderCount: undefined,
		} as InspectorDeckState);
	};
	return { inspector, select };
}

const section = (id: string) =>
	document.querySelector<SectionHost>(`[data-pptx-format-section="${id}"]`)!;
const shownIds = () =>
	[...document.querySelectorAll<HTMLElement>('[data-pptx-format-section]')]
		.filter((node) => !node.hidden)
		.map((node) => node.dataset.pptxFormatSection);

describe('vanilla Format pane sections', () => {
	it('orders a shape as Fill, Outline open, Effects closed, with panels inside', () => {
		const { select } = mount(createFormatPaneController());
		select(shape('a'));
		expect(shownIds().slice(0, 3)).toEqual(['fill', 'outline', 'effects']);
		expect(section('fill').open).toBe(true);
		expect(section('outline').open).toBe(true);
		expect(section('effects').open).toBe(false);
		expect(section('sizePosition').open).toBe(false);
		expect(
			section('fill').querySelector('[data-pptx-inspector-section="fill-stroke"]'),
		).not.toBeNull();
		expect(
			section('sizePosition').querySelector('[data-pptx-inspector-section="transform"]'),
		).not.toBeNull();
		expect(section('slide').hidden).toBe(true);
	});

	it('remembers a toggled section for the next element of the same category', () => {
		const pane = createFormatPaneController();
		const { select } = mount(pane);
		select(shape('a'));
		section('fill').shadowRoot!.querySelector('summary')!.click();
		expect(section('fill').open).toBe(false);
		select(shape('b'));
		expect(section('fill').open).toBe(false);
		expect(pane.isOpen('shape', { id: 'fill', labelKey: '', defaultOpen: true })).toBe(false);
	});

	it('expands and focuses a requested section', () => {
		vi.useFakeTimers({ toFake: ['requestAnimationFrame'] });
		try {
			const pane = createFormatPaneController();
			const { select } = mount(pane);
			select(shape('a'));
			pane.requestFormatPane('sizePosition', { focus: true });
			expect(section('sizePosition').open).toBe(true);
			vi.advanceTimersToNextFrame();
			const summary = section('sizePosition').shadowRoot!.querySelector('summary');
			expect(section('sizePosition').shadowRoot!.activeElement).toBe(summary);
		} finally {
			vi.useRealTimers();
		}
	});

	it('drops Size and position and tags differing values for a multi-selection', () => {
		const { select } = mount(createFormatPaneController());
		const a = shape('a', '#ff0000');
		select(a, [a, shape('b', '#00ff00')]);
		expect(shownIds()).not.toContain('sizePosition');
		expect(section('fill').mixed).toBe(true);
		expect(section('outline').mixed).toBe(false);
	});
});
