import { describe, expect, it } from 'vitest';

import {
	translationsDe,
	translationsEs,
	translationsFr,
	translationsJa,
	translationsKo,
	translationsZhCN,
} from '.';
import { translationsEn } from '../../../../../src/ui/src/pptx/i18n';
import { officeStatusBarState } from '../../../../../src/ui/src/pptx/web-components/status-bar';
import { labelInName } from '../../../../../src/ui/src/test-support/label-in-name';

const locales = {
	en: translationsEn,
	de: translationsDe,
	es: translationsEs,
	fr: translationsFr,
	'zh-CN': translationsZhCN,
	ja: translationsJa,
	ko: translationsKo,
};

describe('status bar names contain the text they show (WCAG 2.5.3)', () => {
	for (const [locale, translations] of Object.entries(locales))
		it(locale, () => {
			const state = officeStatusBarState({
				slideCount: 3,
				activeSlideIndex: 0,
				saveText: '',
				zoomPercent: 100,
				showNotes: true,
				translate: (key) => (translations as Record<string, string>)[key] ?? key,
			});
			const mismatched = (state.toggles ?? [])
				.filter((button) => button.text && !labelInName(button.text, button.label))
				.map((button) => `"${button.text}" -> "${button.label}"`);
			expect(mismatched).toEqual([]);
		});
});
