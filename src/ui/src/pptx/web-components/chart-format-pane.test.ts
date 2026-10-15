// @vitest-environment happy-dom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { registerPptxWebControls } from '.';
import { translationsEn } from '../i18n';
import { columnChart } from '../render/chart-format/chart-format-fixtures.test-util';
import type { ChartFormatChangeEvent, ChartFormatTargetEvent } from './chart-format-pane';

registerPptxWebControls();
afterEach(() => document.body.replaceChildren());

const t = (key: string, params?: Readonly<Record<string, string | number>>) =>
	(translationsEn[key] ?? key).replace(/\{\{(\w+)\}\}/gu, (_, n: string) =>
		String(params?.[n] ?? ''),
	);

function pane(target = 'chartArea') {
	const node = document.createElement('pptx-ui-chart-format-pane');
	node.setAttribute('inline', '');
	node.translateLabel = t;
	node.chartData = columnChart();
	node.target = target;
	document.body.append(node);
	return node;
}

const q = <T extends Element>(node: HTMLElement, selector: string) =>
	node.shadowRoot!.querySelector<T>(selector)!;

describe('pptx-ui-chart-format-pane', () => {
	it('titles itself after the target and lists the chart elements', () => {
		const node = pane('legend');
		expect(q(node, 'h3').textContent).toBe('Format Legend');
		const picker = q<HTMLElement & { value: string }>(node, '[data-chart-format-target]');
		expect(picker.value).toBe('legend');
		expect([...picker.querySelectorAll('option')].map((o) => o.textContent)).toContain(
			'Series "East"',
		);
		expect(
			[...node.shadowRoot!.querySelectorAll('[data-chart-format-tab]')].map((b) => b.textContent),
		).toEqual(['Fill & Line', 'Effects', 'Legend Options']);
	});

	it('emits a whole new chart model for a committed field', () => {
		const node = pane('legend');
		const change = vi.fn();
		node.addEventListener('chart-format-change', change);
		const solid = q<HTMLElement & { checked: boolean }>(node, '[data-chart-format-choice="solid"]');
		solid.checked = true;
		solid.dispatchEvent(new Event('change'));
		expect(change).toHaveBeenCalledOnce();
		const detail = (change.mock.calls[0]![0] as ChartFormatChangeEvent).detail;
		expect(detail.target).toBe('legend');
		expect(detail.chartData.elementFormats?.legend?.fill?.kind).toBe('solid');
		// The colour and transparency rows appear once the fill is solid.
		const transparency = q<HTMLInputElement>(
			node,
			'[data-chart-format-field="fill.transparency"] input',
		);
		transparency.value = '35';
		transparency.dispatchEvent(new Event('change'));
		const last = (change.mock.calls[1]![0] as ChartFormatChangeEvent).detail;
		expect(last.chartData.elementFormats?.legend?.fill?.transparency).toBe(35);
	});

	it('switches tabs, retargets from the picker and closes on Escape', () => {
		const node = pane('series:0');
		q<HTMLButtonElement>(node, '[data-chart-format-tab="options"]').click();
		expect(
			node.shadowRoot!.querySelector('[data-chart-format-field="series.gapWidth"]'),
		).not.toBeNull();
		const retarget = vi.fn();
		node.addEventListener('chart-format-target', retarget);
		const picker = q<HTMLElement & { value: string }>(node, '[data-chart-format-target]');
		picker.value = 'axis:1';
		picker.dispatchEvent(new Event('change'));
		expect((retarget.mock.calls[0]![0] as ChartFormatTargetEvent).detail.target).toBe('axis:1');
		expect(q(node, 'h3').textContent).toBe('Format Axis');
		const close = vi.fn();
		node.addEventListener('chart-format-close', close);
		q(node, '.panel').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(close).toHaveBeenCalledOnce();
	});

	it('disables every control on a read-only deck and emits nothing', () => {
		const node = pane('chartArea');
		node.disabled = true;
		const change = vi.fn();
		node.addEventListener('chart-format-change', change);
		const none = q<HTMLElement & { disabled: boolean }>(node, '[data-chart-format-choice="none"]');
		expect(none.disabled).toBe(true);
		none.dispatchEvent(new Event('change'));
		expect(change).not.toHaveBeenCalled();
	});
});
