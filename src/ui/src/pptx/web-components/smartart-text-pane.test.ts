// @vitest-environment happy-dom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import type { PptxSmartArtData } from 'ooxml-core/pptx';

import { registerPptxWebControls } from './index';
import type { PptxUiSmartArtTextPaneElement, SmartArtTextPaneChange } from './smartart-text-pane';

beforeAll(() => registerPptxWebControls());
afterEach(() => document.body.replaceChildren());

const smartArtData: PptxSmartArtData = {
	resolvedLayoutType: 'list',
	nodes: [
		{ id: 'a', text: 'Alpha', parentId: 'doc' },
		{ id: 'b', text: '', parentId: 'doc' },
	],
	connections: [
		{ sourceId: 'doc', destId: 'a', type: 'parOf', srcOrd: 0 },
		{ sourceId: 'doc', destId: 'b', type: 'parOf', srcOrd: 1 },
	],
};

function mount(state: Partial<PptxUiSmartArtTextPaneElement['state']> = {}) {
	const pane = document.createElement('pptx-ui-smartart-text-pane');
	document.body.append(pane);
	pane.state = { smartArtData, canEdit: true, translate: (key) => `T(${key})`, ...state };
	const seen: SmartArtTextPaneChange[] = [];
	pane.addEventListener('pptx-smartart-data-change', (event) =>
		seen.push((event as CustomEvent<SmartArtTextPaneChange>).detail),
	);
	return { pane, seen };
}
const inputs = (pane: HTMLElement) => [
	...pane.shadowRoot!.querySelectorAll<HTMLInputElement>('input[data-node]'),
];

describe('pptx-ui-smartart-text-pane', () => {
	it('shows the diagram with the pptx labels and tools', () => {
		const { pane } = mount();
		expect(pane.shadowRoot!.querySelector('.title')!.textContent).toBe(
			'T(pptx.smartart.textPane) (2)',
		);
		expect(pane.shadowRoot!.querySelector('.append')!.textContent).toBe('T(pptx.smartart.addItem)');
		expect(inputs(pane).map((input) => input.value)).toEqual(['Alpha', '']);
		expect(pane.shadowRoot!.querySelectorAll('.marker')).toHaveLength(2);
		expect(pane.hidden).toBe(false);
		pane.state = { smartArtData: undefined, canEdit: true };
		expect(pane.hidden).toBe(true);
	});

	it('turns typing, keys and tools into committed SmartArt data', () => {
		const { pane, seen } = mount();
		const [alpha] = inputs(pane) as [HTMLInputElement];
		alpha.value = 'Alpha 2';
		alpha.dispatchEvent(new Event('input', { bubbles: true }));
		expect(seen[0]!.smartArtData.nodes[0]!.text).toBe('Alpha 2');
		// The next request builds on the edit before the host sets state again.
		alpha.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(seen[1]!.smartArtData.nodes.map((n) => n.text)).toEqual(['Alpha 2', '', '']);
		const fill = pane.shadowRoot!.querySelector<HTMLInputElement>('.fill input')!;
		fill.value = '#ff0000';
		fill.dispatchEvent(new Event('change', { bubbles: true }));
		expect(seen[2]!.smartArtData.nodes[0]!.style?.fillColor).toBe('#ff0000');
		expect(seen[2]!.recentColor).toBe('#ff0000');
	});

	it('is read-only without edit rights', () => {
		const { pane, seen } = mount({ canEdit: false });
		const [alpha] = inputs(pane) as [HTMLInputElement];
		expect(alpha.disabled).toBe(true);
		pane.shadowRoot!.querySelector<HTMLButtonElement>('.append')!.click();
		expect(seen).toEqual([]);
	});
});
