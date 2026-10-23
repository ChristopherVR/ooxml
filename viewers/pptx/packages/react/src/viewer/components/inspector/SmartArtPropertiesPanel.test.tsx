// @vitest-environment happy-dom
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSmartArtData } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { renderToStaticMarkup } from 'react-dom/server';
/**
 * Tests for the SmartArt properties panel. The text pane is the shared
 * `pptx-ui-smartart-text-pane` element (its outline, tools and edits are tested
 * in `ooxml-ui`); here the panel must feed it the data and commit what it emits.
 */
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

// Mock react-i18next: return the key so assertions are stable.
vi.mock<typeof import('react-i18next')>(import('react-i18next'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { SmartArtPropertiesPanel } = await import('./SmartArtPropertiesPanel');

function render(el: React.ReactElement): string {
	return renderToStaticMarkup(el);
}

function listData(count: number): PptxSmartArtData {
	return {
		resolvedLayoutType: 'list',
		nodes: Array.from({ length: count }, (_, i) => ({ id: `n${i}`, text: `Item ${i + 1}` })),
	};
}

beforeAll(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	registerPptxWebControls();
});

const mounted: { root: Root; target: HTMLElement }[] = [];
afterEach(() => {
	for (const { root, target } of mounted.splice(0)) {
		act(() => root.unmount());
		target.remove();
	}
});

function mount(data: PptxSmartArtData, canEdit = true) {
	const onUpdateElement = vi.fn<(updates: Partial<PptxElement>) => void>();
	const target = document.createElement('div');
	document.body.append(target);
	const root = createRoot(target);
	mounted.push({ root, target });
	act(() =>
		root.render(
			<SmartArtPropertiesPanel
				smartArtData={data}
				canEdit={canEdit}
				onUpdateElement={onUpdateElement}
			/>,
		),
	);
	const pane = target.querySelector('pptx-ui-smartart-text-pane')!;
	const inputs = () => [
		...pane.shadowRoot!.querySelectorAll<HTMLInputElement>('[data-testid="smartart-node-text"]'),
	];
	return { onUpdateElement, pane, inputs };
}

describe('smartArtPropertiesPanel - shared text pane', () => {
	it('wraps the panel in a labelled group and feeds the pane the nodes', () => {
		const { pane, inputs } = mount(listData(2));
		expect(pane.closest('[role="group"]')?.getAttribute('aria-label')).toBe('pptx.smartart.title');
		expect(inputs().map((input) => input.value)).toEqual(['Item 1', 'Item 2']);
	});

	it('commits the edits the pane emits through onUpdateElement', () => {
		const { onUpdateElement, inputs } = mount(listData(2));
		const second = inputs()[1]!;
		second.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
		expect(onUpdateElement).toHaveBeenCalledOnce();
		const next = onUpdateElement.mock.calls[0]![0].smartArtData as PptxSmartArtData;
		expect(next.nodes.find((node) => node.id === 'n1')?.parentId).toBe('n0');
	});

	it('is read-only when canEdit is false', () => {
		const { onUpdateElement, inputs } = mount(listData(2), false);
		expect(inputs().every((input) => input.disabled)).toBe(true);
		inputs()[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(onUpdateElement).not.toHaveBeenCalled();
	});
});

/**
 * These controls used to print the OOXML wire tokens (`colorful1`, `flat`) as
 * their own captions. The regression guard has two halves: the text must be a
 * dictionary key, and the VALUE must still be the wire token, because changing
 * a value would silently move the control out of parity with the other four
 * bindings and would write a different `dgm:` family into the deck.
 */
describe('smartArtPropertiesPanel - schema tokens are spelled, values are not', () => {
	it('labels the colour-scheme options without changing their values', () => {
		const html = render(
			<SmartArtPropertiesPanel smartArtData={listData(1)} canEdit onUpdateElement={vi.fn()} />,
		);
		for (const scheme of [
			'colorful1',
			'colorful2',
			'colorful3',
			'monochromatic1',
			'monochromatic2',
		]) {
			expect(html).toContain(`value="${scheme}"`);
		}
		expect(html).toContain('pptx.smartart.schemeColorful1');
		expect(html).toContain('pptx.smartart.schemeMonochromatic2');
		expect(html).not.toMatch(/>colorful1</u);
		expect(html).not.toMatch(/>monochromatic1</u);
	});

	it('still offers exactly five colour schemes', () => {
		const html = render(
			<SmartArtPropertiesPanel smartArtData={listData(1)} canEdit onUpdateElement={vi.fn()} />,
		);
		expect(html.match(/<option /gu)).toHaveLength(5);
	});

	it('labels the three style buttons without changing what they set', () => {
		const html = render(
			<SmartArtPropertiesPanel smartArtData={listData(1)} canEdit onUpdateElement={vi.fn()} />,
		);
		expect(html).toContain('pptx.smartart.styleFlat');
		expect(html).toContain('pptx.smartart.styleModerate');
		expect(html).toContain('pptx.smartart.styleIntense');
		expect(html).not.toMatch(/>flat</u);
		expect(html).not.toMatch(/>intense</u);
		// The group keeps its role and accessible name.
		expect(html).toContain('aria-label="pptx.smartart.style"');
	});
});
