import { describe, expect, it } from 'vitest';

import { bridgeCss } from './office-token-bridge';

/** Every rule selector in a stylesheet, with `@media` wrappers removed. */
function selectors(css: string): string[] {
	const flat = css.replace(/@media[^{]*\{([^{}]*\{[^}]*\})\s*\}/g, '$1');
	return [...flat.matchAll(/([^{}]+)\{[^}]*\}/g)].map((match) => match[1]!.trim());
}

describe('pptx chrome bridges', () => {
	it.each(['pptx-ui-title-bar', 'pptx-ui-status-bar'] as const)(
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
		expect(css).toContain('--office-title-bar-name-foreground: var(--pptx-foreground, #f1f5f9);');
		expect(css).toContain('--office-title-bar-result-selected: var(--pptx-accent, #33334d);');
		expect(css).toContain('--pptx-switch-width: 28px;');
		expect(css).toContain('--pptx-switch-knob-travel: 13px;');
	});

	it('themes the status bar through the shared tokens and metrics', () => {
		const css = bridgeCss('pptx-ui-status-bar');
		expect(css).toContain('--office-status-bar-height: 29px;');
		expect(css).toContain('--office-status-bar-separator-opacity: 0.6;');
	});
});
