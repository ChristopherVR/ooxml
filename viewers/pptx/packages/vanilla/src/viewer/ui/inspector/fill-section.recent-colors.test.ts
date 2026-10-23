/**
 * The Fill & Line pickers add to the deck's "Recent Colors" (`p:clrMru`) once
 * per COMMITTED pick. A native colour input fires `input` on every shade a drag
 * through the OS picker passes and `change` once when the pick settles; the
 * fill and outline used to be recorded on every `input`, flooding the capped
 * list. Wired end to end here: the section, the chrome's inspector handlers and
 * the real edit actions over a store.
 */
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { buildChromeCallbacks } from '../../chrome-callbacks';
import type { ChromeCallbackDeps } from '../../chrome-callbacks';
import { createEditActions } from '../../editor/editor-edit-ops';
import { createEditorOps } from '../../editor/editor-operations';
import { currentRecentColors } from '../../editor/editor-recent-colors';
import { createTranslator } from '../../i18n';
import { createInitialViewerState, createStore } from '../../state';
import { themeSwatch } from '../theme-color-swatch-grid.test-support';
import { createFillSection } from './fill-section';
import type { InspectorState } from './types';

const THEME: Record<string, string> = {
	dk1: '#000000',
	lt1: '#ffffff',
	dk2: '#44546a',
	lt2: '#e7e6e6',
	accent1: '#4472c4',
	accent2: '#ed7d31',
	accent3: '#a5a5a5',
	accent4: '#ffc000',
	accent5: '#5b9bd5',
	accent6: '#70ad47',
};

const DRAG = ['#aa0000', '#bb0000', '#cc0000'];

type Styled = PptxElement & { shapeStyle: { fillColor?: string; strokeColor?: string } };

function mount(mruColors: string[] = []) {
	const slide = {
		id: 's',
		rId: 'rId1',
		slideNumber: 1,
		notes: '',
		elements: [
			{
				type: 'shape',
				id: 'a',
				x: 0,
				y: 0,
				width: 100,
				height: 50,
				shapeType: 'rect',
				shapeStyle: { fillColor: '#ffffff', strokeColor: '#000000', strokeWidth: 1 },
			} as PptxElement,
		],
	} as PptxSlide;
	const store = createStore({
		...createInitialViewerState(),
		slides: [slide],
		currentSlide: 0,
		editable: true,
		selectedElementId: 'a',
		presentationProperties: { mruColors },
	});
	const ops = createEditorOps({ store, getHandler: () => null, onHistoryChange: vi.fn() });
	const actions = createEditActions({
		doc: document,
		getTranslator: () => createTranslator(),
		store,
		ops,
		getHandler: () => null,
		setHandler: () => {},
	});
	// The viewer's own inspector handlers, so their forwarding is covered too.
	const deps = new Proxy({ getEditActions: () => actions } as unknown as ChromeCallbackDeps, {
		get: (target, property, receiver) => Reflect.get(target, property, receiver) ?? vi.fn(),
	});
	const section = createFillSection(
		document,
		(key) => key,
		() => document.createElement('section'),
		buildChromeCallbacks(deps).inspectorHandlers,
	);
	section.update({
		hasSelection: true,
		canShape: true,
		fillColor: '#ffffff',
		strokeColor: '#000000',
		strokeWidth: 1,
		fillOpacity: 1,
		strokeOpacity: 1,
		gradientEnabled: false,
		gradient: { type: 'linear', angle: 90, stops: [] },
		themeColorMap: THEME,
		recentColors: currentRecentColors(store.get()),
	} as unknown as InspectorState);
	const [fill, stroke] = Array.from(
		section.el.querySelectorAll<HTMLInputElement>(':scope > [role="group"] input[type="color"]'),
	);
	const groups = Array.from(section.el.querySelectorAll<HTMLElement>(':scope > [role="group"]'));
	return {
		fill,
		stroke,
		groups,
		recent: () => store.get().presentationProperties.mruColors ?? [],
		shape: () => store.get().slides[0].elements[0] as Styled,
	};
}

function fire(input: HTMLInputElement, type: 'input' | 'change', value: string): void {
	input.value = value;
	input.dispatchEvent(new Event(type, { bubbles: true }));
}

/** A recent swatch by colour, in the row itself or inside a shared grid's shadow root. */
function recentSwatch(group: HTMLElement, hex: string): HTMLButtonElement | undefined {
	const row = group.querySelector<HTMLElement>('[data-testid="pptx-color-recent"]');
	const buttons = [
		...(row?.querySelectorAll<HTMLButtonElement>('button') ?? []),
		...(row?.shadowRoot?.querySelectorAll<HTMLButtonElement>('button') ?? []),
	];
	return buttons.find(
		(button) => button.getAttribute('aria-label')?.toLowerCase() === `recent ${hex}`,
	);
}

describe('fill section recent colours', () => {
	it.each([
		['fill', 'fillColor'],
		['line', 'strokeColor'],
	] as const)(
		'a drag through the %s picker applies every shade but records only the settled one',
		(which, key) => {
			const view = mount();
			const input = which === 'fill' ? view.fill : view.stroke;

			for (const shade of DRAG) {
				fire(input, 'input', shade);
				expect(view.shape().shapeStyle[key]).toBe(shade);
			}
			expect(view.recent()).toStrictEqual([]);

			fire(input, 'change', '#cc0000');
			expect(view.recent()).toStrictEqual(['#CC0000']);
		},
	);

	it('a recent swatch click still moves that colour to the front', () => {
		const view = mount(['#112233', '#445566']);
		const swatch = recentSwatch(view.groups[0], '#445566');
		expect(swatch).toBeDefined();
		swatch!.click();
		expect(view.shape().shapeStyle.fillColor).toBe('#445566');
		expect(view.recent()).toStrictEqual(['#445566', '#112233']);
	});

	it('a theme swatch click still records the resolved colour', () => {
		const view = mount();
		const swatch = themeSwatch(view.groups[1], 'Accent 2');
		expect(swatch).not.toBeNull();
		swatch!.click();
		expect(view.shape().shapeStyle.strokeColor).toBe('#ed7d31');
		expect(view.recent()).toStrictEqual(['#ED7D31']);
	});
});
