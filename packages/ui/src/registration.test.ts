import { OFFICE_UI_TAGS, registerOfficeUi, THEME_CSS, installOfficeUiTheme } from './index.js';
import { CONTRACT_REVISION, assertContract, defineOnce } from './registry.js';

describe('registration', () => {
	it('defines every tag, idempotently, and installs the theme once', () => {
		registerOfficeUi();
		registerOfficeUi();
		for (const tag of OFFICE_UI_TAGS) expect(customElements.get(tag), tag).toBeDefined();
		expect(document.querySelectorAll('#office-ui-theme')).toHaveLength(1);
		expect(installOfficeUiTheme()).toBe(true);
		expect(document.querySelectorAll('#office-ui-theme')).toHaveLength(1);
	});

	it('uses only the office-ui- prefix', () => {
		for (const tag of OFFICE_UI_TAGS) expect(tag).toMatch(/^office-ui-[a-z-]+$/);
	});

	it('can skip the theme', () => {
		document.getElementById('office-ui-theme')?.remove();
		registerOfficeUi({ theme: false });
		expect(document.getElementById('office-ui-theme')).toBeNull();
	});

	it('rejects a tag stamped by an incompatible contract', () => {
		class Foreign extends HTMLElement {}
		Object.defineProperty(Foreign, Symbol.for('office-ui.web-control-contract'), {
			value: CONTRACT_REVISION + 1,
		});
		customElements.define('office-ui-foreign-contract', Foreign);
		expect(() => assertContract(customElements, 'office-ui-foreign-contract')).toThrow(
			/Incompatible/,
		);
		expect(() => defineOnce(customElements, 'office-ui-foreign-contract', () => Foreign)).toThrow();
	});

	it('does not redefine a tag that is already registered', () => {
		const before = customElements.get('office-ui-button');
		class Other extends HTMLElement {}
		defineOnce(customElements, 'office-ui-button', () => Other);
		expect(customElements.get('office-ui-button')).toBe(before);
	});
});

describe('theme tokens', () => {
	it('carries forced-colors, dark-mode and touch-target rules', () => {
		expect(THEME_CSS).toContain('@media (forced-colors: active)');
		expect(THEME_CSS).toContain('prefers-color-scheme: dark');
		expect(THEME_CSS).toContain('@media (pointer: coarse)');
		expect(THEME_CSS).toContain('--office-target-size: 44px');
		expect(THEME_CSS).toMatch(/Highlight/);
	});
});

describe('shadow styles', () => {
	it('every control ships forced-colors and coarse-pointer rules', () => {
		registerOfficeUi();
		const css = (tag: string): string => {
			const el = document.createElement(tag);
			document.body.append(el);
			const sheets = (el.shadowRoot as ShadowRoot & { adoptedStyleSheets?: CSSStyleSheet[] })
				.adoptedStyleSheets;
			const text = [...el.shadowRoot!.querySelectorAll('style')].map((s) => s.textContent).join('');
			// jsdom has no constructable stylesheets, so the fallback `<style>` carries the CSS.
			return sheets?.length
				? [...sheets].map((s) => [...s.cssRules].map((r) => r.cssText).join('')).join('')
				: text;
		};
		for (const tag of OFFICE_UI_TAGS) {
			if (tag === 'office-ui-smartart' || tag === 'office-ui-icon') continue;
			const text = css(tag);
			expect(text, tag).toContain('forced-colors: active');
		}
		for (const tag of ['office-ui-button', 'office-ui-checkbox', 'office-ui-select']) {
			expect(css(tag), tag).toMatch(/pointer: coarse/);
		}
	});
});
