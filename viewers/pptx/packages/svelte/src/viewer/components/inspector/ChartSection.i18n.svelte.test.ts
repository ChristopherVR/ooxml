import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { I18N_CONTEXT_KEY } from '../../../i18n/context';
import type { Translator } from '../../../i18n/translator';
import { EditorState } from '../../editor/editor-state.svelte';
import ChartSection from './ChartSection.svelte';

/**
 * The chart section used to hard-code its captions in English ("Chart type",
 * "Legend", "Series options", "Axis styling", ...) where every other binding
 * spells them through the translator. Mounted with a marking translator,
 * every caption, summary, aria-label and placeholder must come back marked.
 */
let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

const mark: Translator = (key) => `<${key}>`;

const CHART = {
	type: 'chart',
	id: 'chart-1',
	x: 0,
	y: 0,
	width: 300,
	height: 200,
	chartData: {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [
			{ name: 'Revenue', values: [10, 12], marker: { symbol: 'circle', size: 6 } },
			{ name: 'Cost', values: [4, 5] },
		],
		style: { hasDataLabels: true, hasLegend: true },
		axes: [
			{ axisType: 'catAx', id: 1, position: 'b', majorGridlines: true },
			{ axisType: 'valAx', id: 2, position: 'l' },
		],
	},
} as unknown as PptxElement;

function render(): HTMLElement {
	const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
	editor.editable = true;
	editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [CHART] }]);
	editor.select(CHART.id);
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ChartSection, {
		target,
		props: { editor },
		context: new Map([[I18N_CONTEXT_KEY, mark]]),
	});
	cleanup = () => {
		void unmount(instance);
		target.remove();
	};
	flushSync();
	for (const details of target.querySelectorAll('details')) {
		details.open = true;
	}
	flushSync();
	return target;
}

/** Captions written by the component itself, not chart data (series and category names, numbers). */
const DATA = new Set(['Revenue', 'Cost', 'Q1', 'Q2', '×', '+', '-', '−']);

describe('svelte chart section translation', () => {
	it('spells every caption, aria-label and placeholder through the translator', () => {
		const target = render();
		const english: string[] = [];
		const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
		for (let node = walker.nextNode(); node; node = walker.nextNode()) {
			const text = node.textContent?.trim() ?? '';
			if (/[A-Za-z]{2,}/u.test(text) && !text.includes('<pptx.') && !DATA.has(text)) {
				english.push(text);
			}
		}
		for (const element of target.querySelectorAll('[aria-label], [placeholder], [title]')) {
			for (const name of ['aria-label', 'placeholder', 'title']) {
				const value = element.getAttribute(name);
				if (value && /[A-Za-z]{2,}/u.test(value) && !value.includes('<pptx.') && !DATA.has(value)) {
					english.push(`${name}=${value}`);
				}
			}
		}

		expect(english).toStrictEqual([]);
	});

	it('puts Title before Chart type and offers Grouping for a bar chart, as the other bindings do', () => {
		const target = render();
		const captions = [...target.querySelectorAll('label')].map(
			(label) => label.firstChild?.textContent?.trim() ?? '',
		);

		expect(captions.slice(0, 3)).toStrictEqual([
			'<pptx.chart.title>',
			'<pptx.chart.type>',
			'<pptx.chart.grouping>',
		]);
	});
});
