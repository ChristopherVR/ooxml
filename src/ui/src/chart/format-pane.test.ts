import { applyChartEditIntent } from 'ooxml-core/chart';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { defineChartFormatPane, type ChartFormatEditEvent } from './format-pane';
import { columnSpace } from './format-pane-fixture.test-util';
import { chartSpaceFormatSource } from './format-pane-space';

defineChartFormatPane();
afterEach(() => document.body.replaceChildren());

function pane(target = 'chartArea') {
	const node = document.createElement('office-ui-chart-format-pane');
	node.chartSpace = columnSpace();
	node.target = target;
	document.body.append(node);
	return node;
}

const q = <T extends Element>(node: HTMLElement, selector: string) =>
	node.shadowRoot!.querySelector<T>(selector)!;

describe('office-ui-chart-format-pane', () => {
	it('titles itself after the target, lists the elements and offers the neutral tabs', () => {
		const node = pane('series:0');
		expect(q(node, 'h3').textContent).toBe('Format Data Series');
		expect(node.getAttribute('aria-label')).toBe('Format Data Series');
		expect(node.getAttribute('role')).toBe('complementary');
		const picker = q<HTMLSelectElement>(node, '[data-chart-format-target]');
		expect(picker.value).toBe('series:0');
		expect([...picker.options].map((o) => o.textContent)).toContain('Series "West"');
		expect(
			[...node.shadowRoot!.querySelectorAll('[data-chart-format-tab]')].map((b) => b.textContent),
		).toEqual(['Fill & Line', 'Effects', 'Series Options']);
		expect(
			node.shadowRoot!.querySelector('[data-chart-format-text-options]')?.hasAttribute('disabled'),
		).toBe(true);
	});

	it('fires the edit with its intent and repaints from the model the host sets back', () => {
		const node = pane('legend');
		const edits: ChartFormatEditEvent['detail'][] = [];
		node.addEventListener('office-chart-format-edit', (event) => {
			const detail = (event as ChartFormatEditEvent).detail;
			edits.push(detail);
			if (detail.intent) node.chartSpace = applyChartEditIntent(node.chartSpace!, detail.intent);
		});
		const solid = q<HTMLElement & { checked: boolean }>(node, '[data-chart-format-choice="solid"]');
		solid.checked = true;
		solid.dispatchEvent(new Event('change'));
		expect(edits[0]).toMatchObject({ target: 'legend', field: 'fill.kind', value: 'solid' });
		expect(edits[0]!.intent).toMatchObject({ op: 'fill', target: { kind: 'legend' } });
		const color = q<HTMLInputElement>(node, '[data-chart-format-field="fill.color"] input');
		expect(color.value).toBe('#4472c4');
		color.value = '#112233';
		color.dispatchEvent(new Event('change'));
		expect(node.chartSpace!.legend?.spPr?.fill).toMatchObject({
			kind: 'solid',
			color: { value: '112233' },
		});
	});

	it('shows the model value again when the host refuses an edit', () => {
		const node = pane('series:0');
		q<HTMLButtonElement>(node, '[data-chart-format-tab="options"]').click();
		const gap = q<HTMLInputElement>(node, '[data-chart-format-field="series.gapWidth"] input');
		expect(gap.value).toBe('219');
		gap.value = '900';
		gap.dispatchEvent(new Event('change'));
		expect(gap.value).toBe('219');
	});

	it('draws sliders that preview in the box and commit on change', () => {
		const node = document.createElement('office-ui-chart-format-pane');
		const source = chartSpaceFormatSource(columnSpace());
		node.source = {
			...source,
			flat: true,
			seriesOptions: (index) => ({ ...source.seriesOptions!(index)!, sliders: true }),
		};
		node.target = 'series:0';
		document.body.append(node);
		const change = vi.fn();
		node.addEventListener('office-chart-format-edit', change);
		const slider = q<HTMLInputElement>(node, 'input[type=range][aria-label="Gap Width"]');
		slider.value = '40';
		slider.dispatchEvent(new Event('input'));
		expect(
			q<HTMLInputElement>(node, '[data-chart-format-field="series.gapWidth"] input').value,
		).toBe('40');
		expect(change).not.toHaveBeenCalled();
		slider.dispatchEvent(new Event('change'));
		expect((change.mock.calls[0]![0] as ChartFormatEditEvent).detail.intent).toEqual({
			op: 'groupOptions',
			seriesIndex: 0,
			edit: { gapWidth: 40 },
		});
	});

	it('leaves a slotted section to the host and retargets, closes and locks', () => {
		const node = document.createElement('office-ui-chart-format-pane');
		node.source = { ...chartSpaceFormatSource(columnSpace()), slots: { fill: 'fill' } };
		const own = document.createElement('div');
		own.slot = 'fill';
		node.append(own);
		document.body.append(node);
		expect(q<HTMLSlotElement>(node, 'slot[name="fill"]').assignedElements()).toEqual([own]);
		const retarget = vi.fn();
		node.addEventListener('office-chart-format-target', retarget);
		const picker = q<HTMLSelectElement>(node, '[data-chart-format-target]');
		picker.value = 'plotArea';
		picker.dispatchEvent(new Event('change'));
		expect(retarget.mock.calls[0]![0].detail).toEqual({ target: 'plotArea' });
		expect(q(node, 'h3').textContent).toBe('Format Plot Area');
		const close = vi.fn();
		node.addEventListener('office-chart-format-close', close);
		q(node, '.panel').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(close).toHaveBeenCalledOnce();
		node.disabled = true;
		const edit = vi.fn();
		node.addEventListener('office-chart-format-edit', edit);
		const none = q<HTMLElement & { disabled: boolean }>(node, '[data-chart-format-choice="none"]');
		expect(none.disabled).toBe(true);
		expect(picker.disabled).toBe(true);
		none.dispatchEvent(new Event('change'));
		expect(edit).not.toHaveBeenCalled();
	});

	it('translates captions through translateLabel with the English fallback', () => {
		const node = pane('legend');
		node.translateLabel = (key, _params, fallback) =>
			key === 'chartFormat.title.legend' ? 'Légende' : `${fallback ?? key}!`;
		expect(q(node, 'h3').textContent).toBe('Légende');
		expect(q(node, '[data-chart-format-close]').getAttribute('aria-label')).toBe('Close!');
	});
});
