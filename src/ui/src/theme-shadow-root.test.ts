// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it } from 'vitest';
import {
	installOfficeUiTheme,
	installOfficeUiThemeFor,
	prepareOfficeUiBlock,
	THEME_CSS,
	THEME_STYLE_ID,
} from './theme';
const hosts: HTMLElement[] = [];
const host = () => {
	const node = document.createElement('div');
	document.body.append(node);
	hosts.push(node);
	return node;
};
let existingTheme: HTMLElement | null;
beforeEach(() => {
	existingTheme = document.getElementById(THEME_STYLE_ID);
});

afterEach(async () => {
	await Promise.resolve();
	hosts.splice(0).forEach((node) => node.remove());
	const current = document.getElementById(THEME_STYLE_ID);
	if (current !== existingTheme) current?.remove();
});
it('installs exactly one stylesheet in each embedding shadow root and its owner document', () => {
	const outer = host().attachShadow({ mode: 'open' });
	const innerHost = document.createElement('div');
	outer.append(innerHost);
	const inner = innerHost.attachShadow({ mode: 'open' });
	const toolbar = document.createElement('div');
	inner.append(toolbar);
	expect(installOfficeUiThemeFor(toolbar)).toBe(true);
	installOfficeUiThemeFor(toolbar);
	for (const root of [document, outer, inner]) {
		expect(root.querySelectorAll(`#${THEME_STYLE_ID}`)).toHaveLength(1);
		if (root === document && existingTheme)
			expect(root.getElementById(THEME_STYLE_ID)).toBe(existingTheme);
		else expect(root.getElementById(THEME_STYLE_ID)?.textContent).toBe(THEME_CSS);
	}
});
it.each([false, true])(
	'preserves the owner document when installing in a shadow root (global theme preinstalled: %s)',
	(preinstall) => {
		if (preinstall) installOfficeUiTheme(document);
		const documentTheme = document.getElementById(THEME_STYLE_ID);
		const root = host().attachShadow({ mode: 'open' });
		expect(installOfficeUiTheme(root)).toBe(true);
		installOfficeUiTheme(root);
		expect(root.querySelectorAll('style')).toHaveLength(1);
		expect(document.getElementById(THEME_STYLE_ID)).toBe(documentTheme);
	},
);
it.each(['native control', 'canvas'] as const)(
	'prepares a %s before framework insertion into a shadow root',
	async (kind) => {
		const root = host().attachShadow({ mode: 'open' });
		const toolbar = document.createElement('div');
		if (kind === 'native control') prepareOfficeUiBlock(toolbar, 'toolbar', 'docx');
		else installOfficeUiThemeFor(toolbar);
		expect(root.getElementById(THEME_STYLE_ID)).toBeNull();
		root.append(toolbar);
		await Promise.resolve();
		expect(root.getElementById(THEME_STYLE_ID)).not.toBeNull();
		if (kind === 'native control') expect(toolbar.getAttribute('data-office-product')).toBe('docx');
	},
);
it('inherits touch target geometry through color mode scopes', () => {
	expect(THEME_CSS).toContain(
		':where(:root) { --office-target-size: var(--office-target-size-touch); }',
	);
	expect(THEME_CSS).not.toContain(':where(:root, [data-office-theme]) { --office-target-size:');
});
