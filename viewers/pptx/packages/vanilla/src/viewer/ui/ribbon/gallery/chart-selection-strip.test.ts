import type { RibbonGalleryContext, ToolbarActionId } from 'ooxml-ui/pptx';
import { EMPTY_RIBBON_TRANSITION_DRAFT } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import { createRibbon } from '../ribbon';
import type { RibbonHandlers, RibbonSelectionState } from '../ribbon-types';

/**
 * Chart Format > Current Selection in the vanilla ribbon: the contextual pane
 * draws the shared strip, and its requests run the chart format edit
 * actions.
 */

/** Every method access returns a memoised `vi.fn()`. */
function fakeActions<T extends object>(overrides: Partial<T> = {}): T {
	const cache = new Map<string, ReturnType<typeof vi.fn>>();
	return new Proxy({} as T, {
		get(_target, prop) {
			if (typeof prop !== 'string') {
				return undefined;
			}
			// Optional data host, not a command: a stub function would read as a comments host.
			if (prop === 'reviewComments') {
				return undefined;
			}
			if (prop in overrides) {
				return overrides[prop as keyof T];
			}
			let fn = cache.get(prop);
			if (!fn) {
				fn = vi.fn();
				cache.set(prop, fn);
			}
			return fn;
		},
	});
}

function handlers(): RibbonHandlers {
	return {
		nav: fakeActions(),
		primary: fakeActions(),
		file: fakeActions<RibbonHandlers['file']>({ getRecentPresentationsCount: () => 10 }),
		slideShow: fakeActions<RibbonHandlers['slideShow']>({ showOptions: () => ({}) }),
		insert: fakeActions(),
		edit: fakeActions(),
		findReplace: fakeActions(),
		design: fakeActions(),
		transitions: fakeActions<RibbonHandlers['transitions']>({
			readDraft: () => ({ ...EMPTY_RIBBON_TRANSITION_DRAFT }),
			readTransition: () => undefined,
		}),
		draw: fakeActions(),
	};
}

const CHART = {
	id: 'chart-1',
	type: 'chart',
	x: 0,
	y: 0,
	width: 400,
	height: 300,
	chartData: {
		chartType: 'bar',
		categories: ['A'],
		series: [{ name: 'East', values: [1] }],
		style: { hasLegend: true },
	},
} as unknown as PptxElement;

function selection(extra: Partial<RibbonSelectionState> = {}): RibbonSelectionState {
	const galleryContext: RibbonGalleryContext = { element: CHART };
	return { hasClipboard: false, slideCount: 1, galleryContext, ...extra };
}

function mount(hiddenActions?: ToolbarActionId[]) {
	const h = handlers();
	const ribbon = createRibbon(document, createTranslator(), h, hiddenActions);
	ribbon.setEditState({ editable: true, canUndo: false, canRedo: false });
	ribbon.setEditable(true);
	ribbon.updateSelection(CHART, selection({ chartFormatTarget: 'legend' }));
	ribbon.el.querySelector<HTMLButtonElement>('[data-ribbon-contextual-tab="chartFormat"]')!.click();
	const strip = ribbon.el.querySelector<HTMLElement>(
		'[data-ribbon-group="chartFormat.currentSelection"] pptx-ui-ribbon-home-chart-selection',
	)!;
	return { ribbon, edit: h.edit, strip };
}

const elementsSelect = (strip: HTMLElement) =>
	strip.querySelector<HTMLElement & { value: string }>(
		'[data-ribbon-control="chartFormat.currentSelection.chartElements"]',
	)!;

describe('chart Format > Current Selection (vanilla)', () => {
	it('draws the shared strip with the chart elements and the current one', () => {
		const { strip } = mount();
		expect(strip).not.toBeNull();
		const select = elementsSelect(strip);
		expect(select.value).toBe('legend');
		expect([...select.querySelectorAll('option')].map((o) => o.value)).toContain('series:0');
	});

	it('routes the element pick, Format Selection and Reset to the edit actions', () => {
		const { strip, edit } = mount();
		const select = elementsSelect(strip);
		select.value = 'plotArea';
		select.dispatchEvent(new Event('change'));
		expect(edit.selectChartFormatTarget).toHaveBeenCalledWith('plotArea');
		strip.querySelector<HTMLButtonElement>('[data-testid="chart-format-selection"]')!.click();
		expect(edit.openChartFormatPane).toHaveBeenCalledOnce();
		strip.querySelector<HTMLButtonElement>('[data-testid="chart-format-reset"]')!.click();
		expect(edit.resetChartFormatSelection).toHaveBeenCalledOnce();
	});
});
