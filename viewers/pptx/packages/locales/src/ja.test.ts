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
import { translationsJa } from './ja';

/** Kana or kanji: every Japanese caption has at least one. */
const JAPANESE = /[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]/u;

describe('Japanese locale', () => {
	it('has a native label in the language catalog', () => {
		expect(LOCALE_CATALOG.find(({ code }) => code === 'ja')?.nativeLabel).toBe('日本語');
	});

	it('uses PowerPoint terminology and keeps file formats and addresses intact', () => {
		const expected = {
			'pptx.statusBar.language': '日本語',
			'pptx.common.ok': 'OK',
			'pptx.ribbon.tab.home': 'ホーム',
			'pptx.ribbon.tab.insert': '挿入',
			'pptx.ribbon.tab.transitions': '画面切り替え',
			'pptx.ribbon.tab.slideShow': 'スライド ショー',
		};
		for (const [key, value] of Object.entries(expected)) {
			expect(translationsJa[key]).toBe(value);
		}
		for (const [key, english] of Object.entries(translationsEn)) {
			if (/^(?:https?:\/\/|wss?:\/\/)/u.test(english) && !english.includes(' ')) {
				expect(translationsJa[key]).toBe(english);
			}
			expect(translationsJa[key]).not.toMatch(/[\u200B-\u200D\uFEFF\u2014]/u);
		}
	});

	it('translates the ribbon tabs and dialog captions into Japanese script', () => {
		const tabs = Object.keys(translationsEn).filter((key) => key.startsWith('pptx.ribbon.tab.'));
		expect(tabs.length).toBeGreaterThan(5);
		for (const key of tabs) {
			expect(translationsJa[key], key).toMatch(JAPANESE);
		}
	});

	for (const { binding, languages, copy } of [
		{ binding: 'React', languages: reactLanguages, copy: reactCopy },
		{ binding: 'Vue', languages: vueLanguages, copy: vueCopy },
		{ binding: 'Angular', languages: angularLanguages, copy: angularCopy },
		{ binding: 'Svelte', languages: svelteLanguages, copy: svelteCopy },
		{ binding: 'Vanilla', languages: vanillaLanguages, copy: vanillaCopy },
	]) {
		it(`${binding} offers Japanese and translates the complete demo shell`, () => {
			expect(languages.find(({ code }) => code === 'ja')?.label).toBe('日本語');
			expect(Object.keys(copy.demoStringsJa).sort()).toStrictEqual(
				Object.keys(copy.demoStringsEn).sort(),
			);
			for (const value of Object.values(copy.demoStringsJa)) {
				expect(value).toMatch(JAPANESE);
			}
		});
	}
});
