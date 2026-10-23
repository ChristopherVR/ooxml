import { mount } from '@vue/test-utils';
import type { PptxElement, PptxSmartArtData, PptxSmartArtNode } from 'pptx-viewer-core';
import { resetSmartArtEditCounter } from 'pptx-viewer-core';
import { beforeEach, describe, expect, it } from 'vitest';

import SmartArtPropertiesPanel from './SmartArtPropertiesPanel.vue';
import { setControlValue } from './test-control-value';

function node(id: string, text: string, parentId?: string): PptxSmartArtNode {
	return { id, text, parentId };
}

function smartArtData(overrides: Partial<PptxSmartArtData> = {}): PptxSmartArtData {
	return {
		nodes: [node('n1', 'First'), node('n2', 'Second'), node('n3', 'Third')],
		resolvedLayoutType: 'list',
		colorScheme: 'colorful1',
		style: 'flat',
		...overrides,
	} as PptxSmartArtData;
}

function smartArtElement(data: PptxSmartArtData = smartArtData()): PptxElement {
	return {
		type: 'smartArt',
		id: 'sa-1',
		x: 0,
		y: 0,
		width: 400,
		height: 300,
		smartArtData: data,
	} as PptxElement;
}

function nonSmartArtElement(): PptxElement {
	return {
		type: 'shape',
		id: 'shape-1',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
	} as PptxElement;
}

/** Read the last emitted `update` patch's smartArtData payload. */
function lastSmartArtData(emitted: unknown): PptxSmartArtData {
	const events = emitted as Array<Array<{ smartArtData?: PptxSmartArtData }>> | undefined;
	if (!events || events.length === 0) {
		throw new Error('no update emitted');
	}
	const patch = events[events.length - 1][0];
	if (!patch.smartArtData) {
		throw new Error('patch has no smartArtData');
	}
	return patch.smartArtData;
}

/** The node inputs inside the shared text pane's shadow root. */
function nodeInputs(wrapper: { element: Element }): HTMLInputElement[] {
	const pane = wrapper.element.querySelector('pptx-ui-smartart-text-pane');
	return [
		...(pane?.shadowRoot?.querySelectorAll<HTMLInputElement>(
			'[data-testid="smartart-node-text"]',
		) ?? []),
	];
}

describe('smartArtPropertiesPanel', () => {
	beforeEach(() => {
		resetSmartArtEditCounter();
	});

	it('shows a muted note for non-smartArt elements', () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: nonSmartArtElement() } });
		expect(wrapper.find('pptx-ui-smartart-text-pane').exists()).toBeFalsy();
		expect(wrapper.text()).toContain('Select a SmartArt graphic');
	});

	it('feeds the shared text pane the nodes and edit rights', () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		expect(nodeInputs(wrapper).map((input) => input.value)).toStrictEqual([
			'First',
			'Second',
			'Third',
		]);
		const readOnly = mount(SmartArtPropertiesPanel, {
			props: { element: smartArtElement(), canEdit: false },
		});
		expect(nodeInputs(readOnly).every((input) => input.disabled)).toBeTruthy();
	});

	it('commits the edits the pane emits as an update patch', () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		const second = nodeInputs(wrapper)[1]!;
		second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
		const next = lastSmartArtData(wrapper.emitted('update'));
		expect(next.nodes.find((n) => n.id === 'n2')?.parentId).toBe('n1');
	});

	it('changing the colour scheme emits smartArtData with the new scheme', async () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		await setControlValue(wrapper.get('[data-testid="smartart-color-scheme"]'), 'monochromatic1');

		const next = lastSmartArtData(wrapper.emitted('update'));
		expect(next.colorScheme).toBe('monochromatic1');
	});

	it('toggling the style emits smartArtData with the new style', async () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		await wrapper.get('[data-testid="smartart-style-intense"]').trigger('click');

		const next = lastSmartArtData(wrapper.emitted('update'));
		expect(next.style).toBe('intense');
	});

	it('switching the layout emits smartArtData with the new resolved layout', async () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		await wrapper.get('[data-testid="smartart-layout-cycle"]').trigger('click');

		const next = lastSmartArtData(wrapper.emitted('update'));
		expect(next.resolvedLayoutType).toBe('cycle');
		// Node content is preserved across a layout switch.
		expect(next.nodes).toHaveLength(3);
	});

	it('does not re-emit when switching to the already-active layout', async () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		await wrapper.get('[data-testid="smartart-layout-list"]').trigger('click');
		expect(wrapper.emitted('update')).toBeUndefined();
	});
});

/**
 * These controls used to print their OOXML wire tokens (`colorful1`, `flat`) as
 * their own captions. Text and value are asserted separately: changing a value
 * would write a different `dgm:` family into the deck and move this panel out of
 * parity with the other bindings, so only the spelling may change.
 */
describe('smartArtPropertiesPanel - schema tokens are spelled, values are not', () => {
	it('labels the colour-scheme options without changing their values', () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });
		const options = wrapper.get('[data-testid="smartart-color-scheme"]').findAll('option');

		expect(options.map((o) => (o.element as HTMLOptionElement).value)).toStrictEqual([
			'colorful1',
			'colorful2',
			'colorful3',
			'monochromatic1',
			'monochromatic2',
		]);
		expect(options.map((o) => o.text())).toStrictEqual([
			'Colourful 1',
			'Colourful 2',
			'Colourful 3',
			'Monochromatic 1',
			'Monochromatic 2',
		]);
	});

	it('labels the style buttons without changing what they set', async () => {
		const wrapper = mount(SmartArtPropertiesPanel, { props: { element: smartArtElement() } });

		expect(wrapper.get('[data-testid="smartart-style-flat"]').text()).toBe('Flat');
		expect(wrapper.get('[data-testid="smartart-style-moderate"]').text()).toBe('Moderate');
		expect(wrapper.get('[data-testid="smartart-style-intense"]').text()).toBe('Intense');

		// The button still emits the wire token, not its caption.
		await wrapper.get('[data-testid="smartart-style-moderate"]').trigger('click');
		const next = lastSmartArtData(wrapper.emitted('update'));
		expect(next.style).toBe('moderate');
	});
});
