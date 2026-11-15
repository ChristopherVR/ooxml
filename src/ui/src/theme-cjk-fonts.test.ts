import { describe, expect, it } from 'vitest';

import { CJK_UI_FONTS, THEME_CSS } from './theme';
import { OFFICE_TOKENS } from './tokens';

describe('CJK UI font fallbacks', () => {
	it('keys each language face list on lang, keeping the Latin stack and default untouched', () => {
		expect(OFFICE_TOKENS['--office-font']).not.toMatch(/Gothic|YaHei|Meiryo/u);
		for (const [lang, faces] of [
			['ja', 'Yu Gothic UI'],
			['ko', 'Malgun Gothic'],
			['zh-CN', 'Microsoft YaHei UI'],
		] as const) {
			expect(CJK_UI_FONTS[lang]).toContain(faces);
			const rule = THEME_CSS.split('\n').find((line) => line.includes(`:root:lang(${lang})`));
			expect(rule, lang).toContain(`--office-font: system-ui, -apple-system, "Segoe UI", `);
			expect(rule, lang).toContain(faces);
			expect(rule, lang).toMatch(/, sans-serif; \}$/u);
		}
	});

	it('never puts a Japanese face in the Korean or Chinese stacks (Han glyph shapes differ)', () => {
		expect(CJK_UI_FONTS.ko).not.toMatch(/Gothic UI|Meiryo|Hiragino|JP/u);
		expect(CJK_UI_FONTS['zh-CN']).not.toMatch(/Gothic|Meiryo|Hiragino|JP|KR/u);
	});

	it('loads no web font', () => {
		expect(THEME_CSS).not.toMatch(/@font-face|@import|url\(/u);
	});
});
