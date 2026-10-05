// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { darkTheme, lightTheme } from './defaults';
import { THEME_KEYS, themeToCssVars } from './css-vars';
import { themeTokenText } from './styles';
import { normalizeThemeMode } from './element';

describe('theme tokens', () => {
	it('defines the docx token set plus the grid tokens in both presets', () => {
		expect(THEME_KEYS).toHaveLength(27);
		for (const preset of [lightTheme, darkTheme])
			for (const key of THEME_KEYS) expect(preset[key], key).toBeTruthy();
	});

	it('keeps Excel green, and white cells in dark mode like the Word editor keeps white paper', () => {
		expect(lightTheme.primary).toBe('#217346');
		expect(lightTheme.selectionBorder).toBe('#217346');
		expect(darkTheme.sheetBg).toBe('#ffffff');
		expect(darkTheme.card).not.toBe(lightTheme.card);
	});

	it('emits kebab-case --xve-* properties and skips empty values', () => {
		expect(themeToCssVars({ headerActiveBg: ' #111 ', ring: '' })).toEqual({
			'--xve-header-active-bg': '#111',
		});
		expect(themeToCssVars(undefined)).toEqual({});
	});

	it('generates light, dark and auto token blocks for the shadow root', () => {
		expect(themeTokenText).toContain(`--xve-background:${lightTheme.background}`);
		expect(themeTokenText).toContain(
			`:host([theme="dark"]){--xve-background:${darkTheme.background}`,
		);
		expect(themeTokenText).toContain(
			'@media (prefers-color-scheme: dark){:host(:not([theme="light"]))',
		);
	});

	it('falls back to auto for unknown modes', () => {
		expect(normalizeThemeMode('dark')).toBe('dark');
		expect(normalizeThemeMode('sepia')).toBe('auto');
	});
});
