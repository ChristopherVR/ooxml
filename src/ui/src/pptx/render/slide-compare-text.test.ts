import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	translationsDe,
	translationsEs,
	translationsFr,
	translationsZhCN,
} from '../../../../../viewers/pptx/packages/locales/src';
import { translationsEn } from '../i18n';
import { buildWaterfallViewModel } from './chart-waterfall-map';
import { computeChartLegendLayout } from './chart-legend-layout';
import { compareSlide } from './slide-compare';
import { describeElementChange } from './slide-compare-text';

const locales: Record<string, Record<string, string>> = {
	en: translationsEn,
	de: translationsDe,
	es: translationsEs,
	fr: translationsFr,
	'zh-CN': translationsZhCN,
};

/** A translator that fills `{{name}}` placeholders and returns the key on a miss. */
function translatorFor(dictionary: Record<string, string>) {
	return (key: string, params?: Record<string, string | number>): string => {
		const value = dictionary[key];
		return value === undefined
			? key
			: value.replace(/\{\{(\w+)\}\}/gu, (_, name: string) => String(params?.[name] ?? ''));
	};
}

const shape = (id: string, x: number, text = ''): PptxElement =>
	({
		id,
		type: 'shape',
		x,
		y: 0,
		width: 100,
		height: 50,
		text,
		textSegments: text ? [{ text, style: {} }] : [],
	}) as unknown as PptxElement;

const base = {
	id: 's1',
	elements: [shape('a', 0, 'Old'), shape('gone', 0)],
	notes: 'one',
} as unknown as PptxSlide;
const compare = {
	id: 's1',
	elements: [shape('a', 50, 'New'), shape('new', 0)],
	notes: 'two',
	backgroundColor: '#FF0000',
} as unknown as PptxSlide;

describe('describeElementChange', () => {
	const changes = compareSlide(base, compare);

	it('covers every kind of change', () => {
		expect(changes.length).toBeGreaterThanOrEqual(6);
		expect(changes.every((change) => change.descriptionKey)).toBe(true);
	});

	it('reproduces the English sentence with the English dictionary', () => {
		const t = translatorFor(translationsEn);
		for (const change of changes) {
			expect(describeElementChange(change, t)).toBe(
				change.description.replace(/\(([^)]*), ([^)]*)\)/gu, '($1, $2)'),
			);
		}
	});

	for (const locale of ['de', 'es', 'fr', 'zh-CN']) {
		it(`${locale} shows no English sentence`, () => {
			const t = translatorFor(locales[locale]!);
			for (const change of changes) {
				const shown = describeElementChange(change, t);
				expect(shown).not.toBe(change.description);
				expect(shown).not.toMatch(
					/Element (added|removed)|Moved from|Resized from|Speaker notes|Text content|Background changed/u,
				);
			}
		});
	}

	it('falls back to English when the dictionary has no key', () => {
		for (const change of changes) {
			expect(describeElementChange(change, (key) => key)).toBe(change.description);
		}
	});
});

describe('waterfall legend labels', () => {
	const chartData = {
		chartType: 'waterfall',
		categories: ['A', 'B'],
		series: [{ name: 'S', values: [1, 2] }],
		style: { hasLegend: true },
	};
	const vm = buildWaterfallViewModel(
		{ id: 'c', type: 'chart', x: 0, y: 0, width: 400, height: 300 } as never,
		chartData as never,
		['A', 'B'],
	);

	it('keeps English without a translator', () => {
		expect(computeChartLegendLayout(vm).map((item) => item.label)).toStrictEqual([
			'Increase',
			'Decrease',
			'Total',
		]);
	});

	for (const locale of ['de', 'es', 'fr', 'zh-CN']) {
		it(`${locale} translates the three bar roles`, () => {
			const t = translatorFor(locales[locale]!);
			const labels = computeChartLegendLayout(vm, (key) => t(key)).map((item) => item.label);
			expect(labels).toHaveLength(3);
			expect(labels).not.toContain('Increase');
			expect(labels).not.toContain('Decrease');
			expect(new Set(labels).size).toBe(3);
		});
	}
});
