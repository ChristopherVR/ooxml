import { ribbonTab, ribbonTabButtons } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createRibbonTabBar } from './ribbon-tab-bar';

describe('ribbon tab row', () => {
	it('draws contextual tabs on the shared tab row and quick actions in its actions slot', () => {
		const onSelect = vi.fn();
		const startRecording = vi.fn();
		const bar = createRibbonTabBar(document, createTranslator(), onSelect, undefined, {
			startRecording,
		});
		expect(bar.el.localName).toBe('pptx-ui-ribbon-tabs');
		bar.setContextualTabs(['shapeFormat']);
		bar.setActive('shapeFormat');
		const contextual = ribbonTab(bar.el, '[data-ribbon-contextual-tab]')!;
		expect(contextual.getAttribute('aria-selected')).toBe('true');
		expect(bar.el.querySelector('.pptxv-tabrow-actions')!.slot).toBe('actions');
		contextual.click();
		bar.el.querySelector<HTMLButtonElement>('.pptxv-tabrow-record')!.click();
		expect(onSelect).toHaveBeenCalledWith('shapeFormat');
		expect(startRecording).toHaveBeenCalledOnce();
		bar.setContextualTabs([]);
		expect(ribbonTab(bar.el, '[data-ribbon-contextual-tab]')).toBeNull();
	});

	it('leaves out the tabs Customize Ribbon hides, but never File', () => {
		const bar = createRibbonTabBar(document, createTranslator(), vi.fn());
		bar.setHiddenTabs(new Set(['file', 'design']));
		const labels = ribbonTabButtons(bar.el).map((tab) => tab.textContent);
		expect(labels).toContain('File');
		expect(labels).not.toContain('Design');
	});

	it('sets ScreenTips from the resolver', () => {
		const bar = createRibbonTabBar(document, createTranslator(), vi.fn());
		bar.applyScreenTips(() => undefined);
		expect(ribbonTab(bar.el)!.hasAttribute('title')).toBe(false);
		bar.applyScreenTips((label) => `${label}!`);
		expect(ribbonTab(bar.el)!.title).toBe('File!');
	});
});
