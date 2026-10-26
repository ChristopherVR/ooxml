// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseChartExSpace } from 'ooxml-core/chart';
import { defineChartExOptions, type OfficeUiChartExOptions } from './chartex-options';
import { chartDataSheetLabels } from './data-sheet-locales';
import { CHART_EX_OPTIONS_STRINGS, chartExOptionsText } from './chartex-options-locales';
const space = () =>
	parseChartExSpace(
		'<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">2</cx:pt><cx:pt idx="1">5</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>',
	).chartExSpace;
afterEach(() => document.body.replaceChildren());
function setup() {
	defineChartExOptions();
	const element = document.createElement('office-ui-chart-ex-options') as OfficeUiChartExOptions;
	document.body.append(element);
	element.space = space();
	return element;
}
describe('shared ChartEx controls', () => {
	it('emits title and family actions without mutating its source', () => {
		const element = setup();
		const listener = vi.fn();
		element.addEventListener('office-chart-ex-edit', listener);
		const input = element.shadowRoot!.querySelector<HTMLInputElement>(
			'[aria-label="Chart title"]',
		)!;
		input.value = 'Changed';
		input.dispatchEvent(new Event('change'));
		expect(listener.mock.calls[0]![0].detail.action).toEqual({ op: 'title', text: 'Changed' });
		expect(element.space?.title).toBeUndefined();
		const connectors = element.shadowRoot!.querySelector<HTMLInputElement>(
			'[aria-label="Connector lines"]',
		)!;
		connectors.checked = false;
		connectors.dispatchEvent(new Event('change'));
		expect(listener.mock.calls[1]![0].detail.action).toEqual({
			op: 'waterfall',
			seriesIndex: 0,
			options: { connectorLines: false },
		});
	});
	it('hides tick labels without asking to hide the whole axis', () => {
		const element = setup();
		const listener = vi.fn();
		element.addEventListener('office-chart-ex-edit', listener);
		const tick = element.shadowRoot!.querySelector<HTMLInputElement>(
			'[aria-label="Show tick labels"]',
		)!;
		tick.checked = false;
		tick.dispatchEvent(new Event('change'));
		expect(listener.mock.calls[0]![0].detail.action).toEqual({
			op: 'tickLabels',
			axisIndex: 0,
			visible: false,
		});
	});
	it('localizes all captions and refuses actions while read-only', () => {
		const element = setup();
		element.locale = 'fr';
		expect(element.shadowRoot!.querySelector('[aria-label="Titre du graphique"]')).not.toBeNull();
		element.disabled = true;
		const listener = vi.fn();
		element.addEventListener('office-chart-ex-edit', listener);
		const input = element.shadowRoot!.querySelector<HTMLInputElement>('input')!;
		expect(input.disabled).toBe(true);
		input.dispatchEvent(new Event('change'));
		expect(listener).not.toHaveBeenCalled();
	});
	it('defines every caption in each supported locale', () => {
		for (const [locale, strings] of Object.entries(CHART_EX_OPTIONS_STRINGS)) {
			expect(strings).toHaveLength(CHART_EX_OPTIONS_STRINGS.en.length);
			expect(strings.every((s) => !!s.trim())).toBe(true);
			if (locale !== 'en')
				expect(chartExOptionsText(locale, 'Chart title')).not.toBe('Chart title');
		}
	});
});

it('keeps every data-sheet translation and interpolation placeholder', () => {
	const english = chartDataSheetLabels('en');
	const placeholders = (text: string) =>
		[...text.matchAll(/\{\w+\}/g)].map((match) => match[0]).sort();
	for (const locale of ['de', 'fr', 'es', 'zh-CN']) {
		const labels = chartDataSheetLabels(locale);
		expect(Object.keys(labels)).toEqual(Object.keys(english));
		for (const key of Object.keys(english) as (keyof typeof english)[])
			expect(placeholders(labels[key])).toEqual(placeholders(english[key]));
	}
});
