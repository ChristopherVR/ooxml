// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { DocxEditorElement, registerDocxEditor } from '../index';
import { themeTokenText } from '../styles';
import { darkTheme, lightTheme } from 'ooxml-core/docx/ui';
import { THEME_KEYS, themeToCssVars } from 'ooxml-core/docx/ui';

describe('theme tokens', () => {
	it('mirrors pptx-viewer semantics: 19 color tokens plus radius, set in both presets', () => {
		expect(THEME_KEYS).toHaveLength(20);
		for (const preset of [lightTheme, darkTheme])
			for (const key of THEME_KEYS) expect(preset[key], key).toBeTruthy();
	});

	it('keeps Word blue as the light primary and distinct surfaces in dark', () => {
		expect(lightTheme.primary).toBe('#185abd');
		expect(darkTheme.background).not.toBe(lightTheme.background);
		expect(darkTheme.card).not.toBe(lightTheme.card);
	});

	it('emits kebab-case --dve-* custom properties and skips empty values', () => {
		const vars = themeToCssVars({ cardForeground: ' #111 ', ring: '', radius: '8px' });
		expect(vars).toEqual({ '--dve-card-foreground': '#111', '--dve-radius': '8px' });
		expect(themeToCssVars(undefined)).toEqual({});
		expect(Object.keys(themeToCssVars(lightTheme))).toHaveLength(THEME_KEYS.length);
	});

	it('inherits canonical tokens and resolves legacy overrides without redefining either on the host', () => {
		expect(themeTokenText.replace(/\s/g, '')).toContain(
			'--_dve-background:var(--dve-background,var(--office-background,',
		);
		expect(themeTokenText.replace(/\s/g, '')).toContain(
			'--_dve-primary:var(--dve-primary,var(--office-accent,',
		);
		expect(themeTokenText).not.toMatch(/--dve-[a-z-]+:/);
		expect(themeTokenText).not.toMatch(/--office-[a-z-]+:/);
		expect(themeTokenText).toContain('color-scheme:inherit');
	});
});

describe('theme element API', () => {
	it('reflects theme to the attribute, defaults to auto and rejects unknown values', () => {
		registerDocxEditor();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		expect(editor.theme).toBe('auto');
		editor.theme = 'dark';
		expect(editor.getAttribute('theme')).toBe('dark');
		expect(editor.getAttribute('data-office-theme')).toBe('dark');
		editor.setAttribute('theme', 'light');
		expect(editor.theme).toBe('light');
		editor.setAttribute('theme', 'sepia');
		expect(editor.theme).toBe('auto');
		expect(editor.hasAttribute('data-office-theme')).toBe(false);
	});

	it('applies themeColors inline and clears stale overrides', () => {
		registerDocxEditor();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.themeColors = { primary: '#ff0000', radius: '2px' };
		expect(editor.style.getPropertyValue('--dve-primary')).toBe('#ff0000');
		editor.themeColors = { ring: '#00ff00' };
		expect(editor.style.getPropertyValue('--dve-primary')).toBe('');
		expect(editor.style.getPropertyValue('--dve-ring')).toBe('#00ff00');
		editor.themeColors = undefined;
		expect(editor.style.getPropertyValue('--dve-ring')).toBe('');
	});
});

it('keeps a host-owned canonical scope when the default theme option is applied', () => {
	registerDocxEditor();
	const editor = document.createElement('docx-editor');
	editor.setAttribute('data-office-theme', 'dark');
	editor.theme = 'auto';
	expect(editor.getAttribute('data-office-theme')).toBe('dark');
	editor.theme = 'light';
	expect(editor.getAttribute('data-office-theme')).toBe('light');
	editor.theme = 'auto';
	expect(editor.hasAttribute('data-office-theme')).toBe(false);
});
