import type { PptxChartData } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createAccessibilitySection } from '../ui/inspector/accessibility-section';
import { createChartSection } from '../ui/inspector/chart-section';
import type { InspectorHandlers, InspectorState } from '../ui/inspector/types';
import { buildViewerCss } from './css';

/**
 * The inspector's bare `<label>caption<control></label>` fields used to have no
 * rule at all, so they took the host page's reset: invisible inputs under the
 * demo's Tailwind preflight, browser defaults (and one run of inline text) in
 * the OOXML Office suite, which ships no reset. These checks run with no host
 * reset, as the suite does.
 */
const mounted: Element[] = [];

function inspector(...sections: HTMLElement[]): HTMLElement {
	const style = document.createElement('style');
	style.textContent = buildViewerCss();
	document.head.appendChild(style);
	const root = document.createElement('div');
	root.className = 'pptxv-inspector';
	root.append(...sections);
	document.body.appendChild(root);
	mounted.push(style, root);
	return root;
}

afterEach(() => {
	for (const node of mounted.splice(0)) {
		node.remove();
	}
});

const section = (): HTMLElement => {
	const el = document.createElement('div');
	el.className = 'pptxv-inspector-section';
	return el;
};

const CHART: PptxChartData = {
	chartType: 'bar',
	categories: ['Q1', 'Q2'],
	series: [
		{ name: 'Revenue', values: [1, 2] },
		{ name: 'Cost', values: [3, 4] },
	],
};

function chartSection(t = (key: string): string => key) {
	const chart = createChartSection(document, t, section, {
		setChartData: vi.fn(),
		getChartFollowDataPoint: () => false,
		pushRecentColor: vi.fn(),
	} as unknown as InspectorHandlers);
	chart.update({ isChart: true, chartData: CHART } as InspectorState);
	return chart;
}

describe('inspector baseline for bare form markup', () => {
	it('lays every chart field out as its own row', () => {
		const chart = chartSection();
		inspector(chart.el);
		const labels = [...chart.el.querySelectorAll<HTMLLabelElement>('label')];
		const title = labels.find((label) => label.textContent?.startsWith('pptx.chart.title'))!;
		const axisMin = chart.el.querySelector<HTMLLabelElement>('.pptxv-exhaustive-field')!;

		expect(getComputedStyle(title).display).toBe('flex');
		expect(getComputedStyle(axisMin).display).toBe('flex');
	});

	it('draws a themed border on a bare text input instead of the host default', () => {
		const chart = chartSection();
		inspector(chart.el);
		const input = chart.el.querySelector<HTMLInputElement>('label > input:not([type])')!;
		const style = getComputedStyle(input);

		expect(style.borderTopStyle).toBe('solid');
		expect(style.boxSizing).toBe('border-box');
		expect(style.backgroundColor).not.toBe('');
	});

	it('resets the action fieldset and gives the alt-text textarea the panel font', () => {
		const fieldset = document.createElement('fieldset');
		const host = section();
		host.append(fieldset);
		const accessibility = createAccessibilitySection(
			document,
			(key) => key,
			section,
			{} as InspectorHandlers,
		);
		inspector(host, accessibility.el);
		const textarea = accessibility.el.querySelector('textarea')!;

		expect(getComputedStyle(fieldset).paddingTop).toBe('0px');
		expect(getComputedStyle(textarea).fontFamily).not.toMatch(/monospace/u);
		expect(getComputedStyle(textarea).display).toBe('block');
	});

	it('leaves a classed label to its own rule', () => {
		const host = section();
		const label = document.createElement('label');
		label.className = 'pptxv-field';
		host.append(label);
		inspector(host);

		expect(getComputedStyle(label).display).toBe('inline-flex');
		expect(getComputedStyle(label).flexWrap).not.toBe('wrap');
	});
});

describe('chart series colour caption', () => {
	it('names the selected series instead of showing the {{name}} placeholder', () => {
		const chart = chartSection(createTranslator());
		const caption = [...chart.el.querySelectorAll('label')].find((label) =>
			label.textContent?.endsWith('colour'),
		)!;

		expect(caption.textContent).toBe('Revenue colour');
		expect(caption.querySelector('input')!.getAttribute('aria-label')).toBe('Revenue colour');
	});
});
