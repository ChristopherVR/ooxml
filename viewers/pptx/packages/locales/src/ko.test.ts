import { describe, expect, it } from 'vitest';

import * as angularCopy from '../../../../../demos/pptx/demo-angular/src/demo-locales';
import { LANGUAGES as angularLanguages } from '../../../../../demos/pptx/demo-angular/src/languages';
import * as reactCopy from '../../../../../demos/pptx/demo-react/demo-locales';
import { languages as reactLanguages } from '../../../../../demos/pptx/demo-react/languages';
import * as svelteCopy from '../../../../../demos/pptx/demo-svelte/src/demo-locales';
import { languages as svelteLanguages } from '../../../../../demos/pptx/demo-svelte/src/languages';
import * as vanillaCopy from '../../../../../demos/pptx/demo-vanilla/src/demo-locales';
import { languages as vanillaLanguages } from '../../../../../demos/pptx/demo-vanilla/src/languages';
import * as vueCopy from '../../../../../demos/pptx/demo-vue/src/demo-locales';
import { languages as vueLanguages } from '../../../../../demos/pptx/demo-vue/src/languages';
import { LOCALE_CATALOG, translationsEn } from '../../../../../src/ui/src/pptx/i18n';
import { translationsKo } from './ko';

/** Hangul: every Korean caption has at least one syllable. */
const KOREAN = /\p{Script=Hangul}/u;

describe('Korean locale', () => {
	it('has a native label in the language catalog', () => {
		expect(LOCALE_CATALOG.find(({ code }) => code === 'ko')?.nativeLabel).toBe('한국어');
	});

	it('uses PowerPoint terminology and keeps file formats and addresses intact', () => {
		const expected = {
			'pptx.statusBar.language': '한국어',
			'pptx.common.ok': '확인',
			'pptx.ribbon.tab.home': '홈',
			'pptx.ribbon.tab.insert': '삽입',
			'pptx.ribbon.tab.transitions': '전환',
			'pptx.ribbon.tab.slideShow': '슬라이드 쇼',
		};
		for (const [key, value] of Object.entries(expected)) {
			expect(translationsKo[key]).toBe(value);
		}
		for (const [key, english] of Object.entries(translationsEn)) {
			if (/^(?:https?:\/\/|wss?:\/\/)/u.test(english) && !english.includes(' ')) {
				expect(translationsKo[key]).toBe(english);
			}
			expect(translationsKo[key]).not.toMatch(/[\u200B-\u200D\uFEFF\u2014]/u);
		}
	});

	it('translates the ribbon tabs and dialog captions into Hangul', () => {
		const tabs = Object.keys(translationsEn).filter((key) => key.startsWith('pptx.ribbon.tab.'));
		expect(tabs.length).toBeGreaterThan(5);
		for (const key of tabs) {
			expect(translationsKo[key], key).toMatch(KOREAN);
		}
	});

	for (const { binding, languages, copy } of [
		{ binding: 'React', languages: reactLanguages, copy: reactCopy },
		{ binding: 'Vue', languages: vueLanguages, copy: vueCopy },
		{ binding: 'Angular', languages: angularLanguages, copy: angularCopy },
		{ binding: 'Svelte', languages: svelteLanguages, copy: svelteCopy },
		{ binding: 'Vanilla', languages: vanillaLanguages, copy: vanillaCopy },
	]) {
		it(`${binding} offers Korean and translates the complete demo shell`, () => {
			expect(languages.find(({ code }) => code === 'ko')?.label).toBe('한국어');
			expect(Object.keys(copy.demoStringsKo).sort()).toStrictEqual(
				Object.keys(copy.demoStringsEn).sort(),
			);
			for (const value of Object.values(copy.demoStringsKo)) {
				expect(value).toMatch(KOREAN);
			}
		});
	}
});
