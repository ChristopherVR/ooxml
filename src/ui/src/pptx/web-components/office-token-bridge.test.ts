import { describe, expect, it } from 'vitest';

import { bridgeCss } from './office-token-bridge';

/** Every rule selector in a stylesheet, with `@media` wrappers removed. */
function selectors(css: string): string[] {
	const flat = css.replace(/\/\*[\s\S]*?\*\//g, '').replace(/@media[^{]*\{/g, '');
	return [...flat.matchAll(/([^{}]+)\{[^}]*\}/g)].map((match) => match[1]!.trim());
}

describe('pptx chrome bridges', () => {
	it.each(['pptx-ui-title-bar', 'pptx-ui-status-bar', 'pptx-ui-color-grid'] as const)(
		'%s styles only its host, never the shared element internals',
		(tag) => {
			const css = bridgeCss(tag);
			expect(selectors(css).length).toBeGreaterThan(0);
			for (const selector of selectors(css)) {
				expect(selector).toMatch(/^:host\b/);
			}
		},
	);

	it('themes the title bar through the shared tokens and metrics', () => {
		const css = bridgeCss('pptx-ui-title-bar');
		expect(css).toContain('--office-title-bar-height: 36px;');
		expect(css).toContain('--office-title-bar-name-foreground: var(--_pptx-foreground, #f1f5f9);');
		expect(css).toContain('--office-title-bar-result-selected: var(--_pptx-accent, #33334d);');
		expect(css).toContain('--pptx-switch-width: 28px;');
		expect(css).toContain('--pptx-switch-knob-travel: 13px;');
	});

	it('keeps the pptx 2px between colour swatches', () => {
		expect(bridgeCss('pptx-ui-color-grid')).toContain('--office-color-swatch-gap: 2px;');
	});

	it('themes the status bar through the shared tokens and metrics', () => {
		const css = bridgeCss('pptx-ui-status-bar');
		expect(css).toContain('--office-status-bar-height: 29px;');
		expect(css).toContain('--office-status-bar-separator-opacity: 0.6;');
	});
});

it('captures canonical inputs on the host and bridges inner chrome without a token cycle', () => {
	const css = bridgeCss('pptx-ui-checkbox');
	expect(css).toContain('--_pptx-background: var(--pptx-background, var(--office-background,');
	expect(css).toContain(':host > * {');
	expect(css).toContain('--office-background: var(--_pptx-background);');
	expect(css).not.toContain('--office-background: var(--office-background');
});
