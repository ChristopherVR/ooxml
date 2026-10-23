import type { PptxSmartArtData } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createSmartArtSection } from './smartart-section';
import type { InspectorHandlers, InspectorState } from './types';

/** A `section()` factory matching the one `createInspector` passes in. */
function sectionFactory() {
	return (): HTMLElement => document.createElement('div');
}

/** Mount with the identity translator so an option's text IS its i18n key. */
function mount() {
	const setSmartArtColorScheme = vi.fn();
	const section = createSmartArtSection(document, (key) => key, sectionFactory(), {
		setSmartArtColorScheme,
		setSmartArtLayout: vi.fn(),
		replaceSmartArtData: vi.fn(),
		pushRecentColor: vi.fn(),
	} as unknown as InspectorHandlers);
	section.update({
		isSmartArt: true,
		smartArtData: { nodes: [], resolvedLayoutType: 'list' } as unknown as PptxSmartArtData,
	} as InspectorState);
	const scheme = section.el.querySelector<HTMLSelectElement>(
		'[data-testid="smartart-color-scheme"]',
	)!;
	return { section, setSmartArtColorScheme, scheme };
}

/** Mount with real nodes and a `replaceSmartArtData` that feeds back through `update`. */
function mountWithNodes(nodes: PptxSmartArtData['nodes'], extra: Partial<InspectorHandlers> = {}) {
	let data: PptxSmartArtData = { nodes, resolvedLayoutType: 'list' } as PptxSmartArtData;
	const replaceSmartArtData = vi.fn((next: PptxSmartArtData) => {
		data = next;
		section.update({ isSmartArt: true, smartArtData: data } as InspectorState);
	});
	const section = createSmartArtSection(document, (key) => key, sectionFactory(), {
		setSmartArtColorScheme: vi.fn(),
		setSmartArtLayout: vi.fn(),
		replaceSmartArtData,
		pushRecentColor: vi.fn(),
		...extra,
	} as unknown as InspectorHandlers);
	document.body.appendChild(section.el);
	section.update({ isSmartArt: true, smartArtData: data } as InspectorState);
	return { section, replaceSmartArtData, getData: () => data };
}

function nodeInputs(el: HTMLElement): HTMLInputElement[] {
	const pane = el.querySelector('pptx-ui-smartart-text-pane');
	return [
		...(pane?.shadowRoot?.querySelectorAll<HTMLInputElement>(
			'[data-testid="smartart-node-text"]',
		) ?? []),
	];
}

describe('smartArtSection shared text pane', () => {
	afterEach(() => {
		document.body.replaceChildren();
	});

	it('feeds the pane the nodes', () => {
		const { section } = mountWithNodes([
			{ id: 'n1', text: 'One' },
			{ id: 'n2', text: 'Two' },
		]);
		expect(nodeInputs(section.el).map((input) => input.value)).toStrictEqual(['One', 'Two']);
	});

	it('commits the structure edits the pane emits', () => {
		const { section, replaceSmartArtData, getData } = mountWithNodes([
			{ id: 'n1', text: 'One' },
			{ id: 'n2', text: 'Two' },
		]);
		nodeInputs(section.el)[1]!.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }),
		);
		expect(replaceSmartArtData).toHaveBeenCalledOnce();
		expect(getData().nodes.find((n) => n.id === 'n2')?.parentId).toBe('n1');
		nodeInputs(section.el)[0]!.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
		);
		expect(getData().nodes).toHaveLength(3);
	});

	it('records a picked node colour as a recent colour', () => {
		const pushRecentColor = vi.fn();
		const { section, getData } = mountWithNodes([{ id: 'n1', text: 'One' }], { pushRecentColor });
		const pane = section.el.querySelector('pptx-ui-smartart-text-pane')!;
		const fill = pane.shadowRoot!.querySelector<HTMLInputElement>('.fill input')!;
		fill.value = '#ff0000';
		fill.dispatchEvent(new Event('change', { bubbles: true }));
		expect(getData().nodes[0]?.style?.fillColor).toBe('#ff0000');
		expect(pushRecentColor).toHaveBeenCalledWith('#ff0000');
	});
});

describe('smartart colour scheme picker', () => {
	it('keeps the five `dgm:colorsDef` families as the option values', () => {
		const { scheme } = mount();

		expect(Array.from(scheme.options).map((option) => option.value)).toStrictEqual([
			'colorful1',
			'colorful2',
			'colorful3',
			'monochromatic1',
			'monochromatic2',
		]);
	});

	it('spells the families rather than showing `monochromatic2`', () => {
		const { scheme } = mount();

		expect(Array.from(scheme.options).map((option) => option.textContent)).toStrictEqual([
			'pptx.smartart.schemeColorful1',
			'pptx.smartart.schemeColorful2',
			'pptx.smartart.schemeColorful3',
			'pptx.smartart.schemeMonochromatic1',
			'pptx.smartart.schemeMonochromatic2',
		]);
	});

	it('still commits the family token', () => {
		const { scheme, setSmartArtColorScheme } = mount();

		scheme.value = 'monochromatic1';
		scheme.dispatchEvent(new Event('change'));

		expect(setSmartArtColorScheme).toHaveBeenCalledWith('monochromatic1');
	});

	it('captions the layout buttons from the shared layout catalogue', () => {
		const { section } = mount();
		const button = section.el.querySelector<HTMLButtonElement>(
			'[data-testid="smartart-layout-hierarchy"]',
		)!;

		expect(button.textContent).toBe('pptx.smartart.category.hierarchy');
	});
});
