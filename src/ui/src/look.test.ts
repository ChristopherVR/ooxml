// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerOfficeUi } from './index';
import {
	applyOfficeLook,
	CARDS_LOOK_CSS,
	CARDS_LOOK_PROPERTIES,
	normalizeOfficeLook,
	OFFICE_LOOK_ATTRIBUTE,
} from './look';
import { LOOK_DARK_TOKENS, LOOK_TOKENS } from './look-tokens';
import { THEME_CSS } from './theme';
import { OFFICE_TOKENS } from './tokens';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function shadowCss(tag: string): string {
	const el = document.createElement(tag);
	document.body.append(el);
	const root = el.shadowRoot!;
	const sheets = (root as ShadowRoot & { adoptedStyleSheets?: CSSStyleSheet[] }).adoptedStyleSheets;
	return sheets?.length
		? sheets.map((s) => [...s.cssRules].map((r) => r.cssText).join('\n')).join('\n')
		: [...root.querySelectorAll('style')].map((s) => s.textContent).join('\n');
}

describe('the cards look', () => {
	it('adds its tokens to the contract with light and dark values', () => {
		for (const [name, value] of Object.entries(LOOK_TOKENS)) {
			expect(OFFICE_TOKENS[name as keyof typeof OFFICE_TOKENS]).toBe(value);
			expect(THEME_CSS, name).toContain(`${name}: ${value};`);
		}
		expect(LOOK_TOKENS['--office-shadow-card']).toBe(
			'0 1px 3px rgb(60 64 67 / 20%), 0 2px 6px rgb(60 64 67 / 8%)',
		);
		expect(LOOK_TOKENS['--office-card-border']).toBe('transparent');
		const dark = THEME_CSS.slice(THEME_CSS.indexOf('[data-office-theme="dark"]'));
		for (const [name, value] of Object.entries(LOOK_DARK_TOKENS))
			expect(dark, name).toContain(`${name}: ${value};`);
	});

	it('declares only private properties, after the mode scopes, and never on :root', () => {
		const names = [...CARDS_LOOK_CSS.matchAll(/(--[\w-]+):/g)].map((match) => match[1]!);
		expect(names.length).toBeGreaterThan(0);
		for (const name of names) expect(name.startsWith('--_office-look-'), name).toBe(true);
		expect(CARDS_LOOK_CSS).not.toContain(':root');
		expect(CARDS_LOOK_CSS).toContain('[data-office-look="cards"] [data-office-theme]');
		expect(THEME_CSS.lastIndexOf(CARDS_LOOK_CSS)).toBeGreaterThan(
			THEME_CSS.lastIndexOf('[data-office-theme="dark"]'),
		);
	});

	it('has no custom-property cycles', () => {
		for (const [name, value] of Object.entries(CARDS_LOOK_PROPERTIES)) {
			const reads = [...value.matchAll(/var\((--[\w-]+)/g)].map((match) => match[1]!);
			expect(reads, name).not.toContain(name);
			for (const read of reads)
				expect(read in CARDS_LOOK_PROPERTIES, `${name} -> ${read}`).toBe(false);
		}
	});

	it('reaches shadow roots through fallbacks that keep the default look', () => {
		expect(shadowCss('office-ui-format-section')).toContain(
			'var(--_office-look-heading-height, var(--office-target-size, 28px))',
		);
		const status = shadowCss('office-ui-status-bar');
		expect(status).toContain('var(--_office-look-pill, var(--office-radius-sm, 3px))');
		expect(status).toContain('var(--_office-look-pill-track, transparent)');
		expect(shadowCss('office-ui-menu-button')).toContain(
			'var(--_office-look-menu-shadow, var(--office-shadow,',
		);
	});

	it('sets and clears the scope attribute without touching anything else', () => {
		const scope = document.createElement('div');
		scope.setAttribute('data-office-theme', 'dark');
		applyOfficeLook(scope, 'cards');
		expect(scope.getAttribute(OFFICE_LOOK_ATTRIBUTE)).toBe('cards');
		applyOfficeLook(scope, 'default');
		expect(scope.hasAttribute(OFFICE_LOOK_ATTRIBUTE)).toBe(false);
		expect(scope.getAttribute('data-office-theme')).toBe('dark');
		expect(normalizeOfficeLook('cards')).toBe('cards');
		expect(normalizeOfficeLook('paper')).toBe('default');
		expect(normalizeOfficeLook(undefined)).toBe('default');
	});
});
