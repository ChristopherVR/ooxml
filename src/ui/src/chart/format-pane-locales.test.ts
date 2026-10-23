import { afterEach, describe, expect, it } from 'vitest';

import { defineChartFormatPane } from './format-pane';
import { columnSpace } from './format-pane-fixture.test-util';
import { CHART_FORMAT_LOCALES, CHART_FORMAT_STRINGS } from './format-pane-locales/index';
import { chartFormatLocale, chartFormatTranslator } from './format-pane-translate';

defineChartFormatPane();
afterEach(() => document.body.replaceChildren());

const placeholders = (text: string) => [...text.matchAll(/\{\{\w+\}\}/g)].map((m) => m[0]).sort();

describe('chart Format pane dictionary', () => {
	const english = Object.keys(CHART_FORMAT_STRINGS.en).sort();

	for (const locale of CHART_FORMAT_LOCALES)
		it(`${locale} has every key, no empty text and the English placeholders`, () => {
			const strings = CHART_FORMAT_STRINGS[locale];
			expect(Object.keys(strings).sort()).toEqual(english);
			for (const key of english) {
				expect(strings[key]?.trim(), key).not.toBe('');
				expect(placeholders(strings[key]!), `${locale}:${key}`).toEqual(
					placeholders(CHART_FORMAT_STRINGS.en[key]!),
				);
			}
		});

	it('maps tags onto its locales and lets a product override a caption', () => {
		expect(chartFormatLocale('fr-CA')).toBe('fr');
		expect(chartFormatLocale('zh-Hans')).toBe('zh-CN');
		expect(chartFormatLocale('ja')).toBe('en');
		const translate = chartFormatTranslator('de', (key) =>
			key === 'chartFormat.close' ? 'Zu' : key,
		);
		expect(translate('chartFormat.close')).toBe('Zu');
		expect(translate('chartFormat.target.chartArea')).toBe(
			CHART_FORMAT_STRINGS.de['chartFormat.target.chartArea'],
		);
		expect(translate('chartFormat.unknown')).toBe('chartFormat.unknown');
	});

	it('covers every caption the pane shows and draws them in its locale', () => {
		const node = document.createElement('office-ui-chart-format-pane');
		node.chartSpace = columnSpace();
		node.locale = 'es';
		const seen = new Set<string>();
		node.translateLabel = (key) => {
			seen.add(key);
			return key;
		};
		document.body.append(node);
		const root = node.shadowRoot!;
		for (const option of [...root.querySelectorAll('[data-chart-format-target] option')]) {
			node.target = (option as HTMLOptionElement).value;
			for (const tab of root.querySelectorAll<HTMLElement>('[data-chart-format-tab]')) tab.click();
		}
		expect(seen.size).toBeGreaterThan(20);
		expect([...seen].filter((key) => !(key in CHART_FORMAT_STRINGS.en))).toEqual([]);
		node.target = 'chartArea';
		expect(root.querySelector('h3')?.textContent).toBe(
			CHART_FORMAT_STRINGS.es['chartFormat.title.chartArea'],
		);
	});
});
