import { afterEach, expect, it } from 'vitest';
import {
	BUILDING_BLOCKS_CSS,
	installOfficeUiTheme,
	prepareOfficeUiBlock,
	THEME_CSS,
	THEME_STYLE_ID,
} from './theme';

afterEach(() => {
	document.getElementById(THEME_STYLE_ID)?.remove();
});

it('installs one theme in the control owner document without replacing live controls', () => {
	const doc = document.implementation.createHTMLDocument('embedded editor');
	const toolbar = doc.createElement('div');
	const button = doc.createElement('button');
	toolbar.append(button);
	doc.body.append(toolbar);
	prepareOfficeUiBlock(toolbar, 'toolbar', 'docx');
	prepareOfficeUiBlock(toolbar, 'tabs', 'xlsx');
	expect(toolbar.firstChild).toBe(button);
	expect(toolbar.dataset.officeBlock).toBe('tabs');
	expect(toolbar.dataset.officeProduct).toBe('xlsx');
	expect(doc.querySelectorAll(`#${THEME_STYLE_ID}`)).toHaveLength(1);
	expect(doc.getElementById(THEME_STYLE_ID)?.textContent).toContain(BUILDING_BLOCKS_CSS);
	expect(document.getElementById(THEME_STYLE_ID)).toBeNull();
});

it('keeps host-owned theme choices and tokens when preparing a standalone control', () => {
	const scope = document.createElement('section');
	scope.dataset.officeTheme = 'dark';
	scope.style.setProperty('--office-accent', '#123456');
	const field = document.createElement('select');
	scope.append(field);
	prepareOfficeUiBlock(field, 'tabs', 'visio');
	expect(scope.dataset.officeTheme).toBe('dark');
	expect(scope.style.getPropertyValue('--office-accent')).toBe('#123456');
	expect(field.hasAttribute('data-office-theme')).toBe(false);
	installOfficeUiTheme();
	expect(document.querySelectorAll(`#${THEME_STYLE_ID}`)).toHaveLength(1);
});

it('scopes palette modes without resetting inherited type and layout tokens', () => {
	const lightScope = THEME_CSS.slice(
		THEME_CSS.indexOf(':where([data-office-theme="light"]'),
		THEME_CSS.indexOf(':where([data-office-theme="dark"]'),
	);
	expect(lightScope).toContain('--office-background: #ffffff');
	expect(lightScope).toContain('--office-field-background: var(--office-background');
	expect(lightScope).not.toContain('--office-font:');
	expect(lightScope).not.toContain('--office-radius:');
	expect(THEME_CSS).toContain('[data-office-theme="auto"]');
	expect(BUILDING_BLOCKS_CSS).toContain('select[data-office-block]');
});
