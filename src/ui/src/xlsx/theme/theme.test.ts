// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { darkTheme, lightTheme } from 'ooxml-core/xlsx/ui';
import { THEME_KEYS, themeToCssVars } from 'ooxml-core/xlsx/ui';
import { editorStyleText, themeTokenText } from './styles';
import { defineXlsxEditor } from '../component';
import { normalizeThemeMode } from 'ooxml-core/xlsx/ui';

describe('theme tokens', () => {
	it('defines the docx token set plus the grid tokens in both presets', () => {
		expect(THEME_KEYS).toHaveLength(28);
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

	it('inherits shared Office theme values without shadow-local defaults', () => {
		// The editor reads the accent ink in the Office palette's scheme, not the host page's.
		expect(themeTokenText).toContain(':host{color-scheme:var(--office-color-scheme, light);');
		expect(editorStyleText).not.toContain('--xve-background:');
		expect(themeTokenText).not.toContain('--office-background:');
		expect(editorStyleText).toContain('.xve-frame {');
		// Bare token reads carry the token's default, so derived tokens resolve where they are read.
		expect(editorStyleText).toContain('var(--xve-card, var(--office-surface, #f3f4f6))');
	});

	it('falls back to auto for unknown modes', () => {
		expect(normalizeThemeMode('dark')).toBe('dark');
		expect(normalizeThemeMode('sepia')).toBe('auto');
	});
});

describe('canonical Office tokens', () => {
	it('reads legacy host overrides before canonical semantic colours', () => {
		expect(editorStyleText).toContain('var(--xve-foreground, var(--office-foreground, #1f2937))');
		expect(themeTokenText).toContain(
			'--xlsx-theme-primary: var(--xve-primary, var(--office-accent, #2563eb))',
		);
		expect(editorStyleText).toContain('--office-accent: var(--xlsx-theme-primary);');
		expect(editorStyleText).toContain('var(--xve-border, var(--office-border, #d1d5db))');
	});
});

it('preserves a host-authored canonical scope when the default auto property is applied', () => {
	defineXlsxEditor();
	const editor = document.createElement('xlsx-editor');
	editor.dataset.officeTheme = 'dark';
	editor.theme = 'auto';
	expect(editor.dataset.officeTheme).toBe('dark');
	editor.theme = 'light';
	expect(editor.dataset.officeTheme).toBe('light');
	editor.theme = 'auto';
	expect(editor.dataset.officeTheme).toBeUndefined();
});
