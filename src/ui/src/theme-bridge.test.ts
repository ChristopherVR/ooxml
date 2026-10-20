import { describe, expect, it } from 'vitest';
import { legacyThemeAliases, legacyThemeBridge, shadcnBridge, themeBridge } from './theme-bridge';

it('resolves legacy compatibility on a separate host from the canonical chrome bridge', () => {
	const aliases = legacyThemeAliases(':host', '--dve-', '--_dve-');
	expect(aliases).toContain(
		'--_dve-background: var(--dve-background, var(--office-background, #ffffff));',
	);
	expect(aliases).toContain('--_dve-radius: var(--dve-radius, var(--office-radius, 4px));');
	expect(aliases).not.toContain('--office-background:');
	const chrome = legacyThemeBridge('.frame', '--_dve-');
	expect(chrome).toContain('.frame {');
	expect(chrome).toContain('--office-background: var(--_dve-background);');
	expect(chrome).toContain('--office-radius: var(--_dve-radius);');
	expect(chrome).not.toContain('--dve-background');
});

describe('themeBridge', () => {
	it('declares each mapped token on the selector and skips empty ones', () => {
		const css = themeBridge(':host', {
			'--office-accent': 'var(--brand)',
			'--office-border': '',
		});
		expect(css.startsWith(':host {\n\t--office-accent: var(--brand);\n')).toBe(true);
		expect(css).not.toContain('--office-border:');
	});

	it('declares the tokens derived from a mapped one again, so they follow the product', () => {
		// The page theme resolves `--office-select-background` on :root; without this a light
		// viewer on a dark system painted its selects with the page's dark background.
		const css = themeBridge(':host > *', { '--office-background': 'var(--surface)' });
		expect(css).toContain('--office-select-background: var(--office-background, #ffffff);');
		expect(css).toContain('--office-field-background: var(--office-background, #ffffff);');
		// Through another derived token too.
		expect(css).toContain('--office-popover: var(--office-background, #ffffff);');
		expect(css).toContain('--office-select-option-active:');
		expect(css).not.toContain('--office-field-border:');
	});
});

describe('shadcnBridge', () => {
	const css = shadcnBridge(':host', '--dve-');

	it('maps the shadcn names a product already has onto the office tokens', () => {
		expect(css).toContain('--office-foreground: var(--dve-foreground);');
		expect(css).toContain('--office-surface: var(--dve-card);');
		expect(css).toContain('--office-selected: var(--dve-accent);');
		expect(css).toContain('--office-accent: var(--dve-primary);');
		expect(css).toContain('--office-danger: var(--dve-destructive);');
	});

	it('falls back from the popover text to the plain foreground', () => {
		expect(css).toContain(
			'--office-popover-foreground: var(--dve-popover-foreground, var(--dve-foreground));',
		);
	});

	it('takes literal fallbacks and extra tokens', () => {
		const withOptions = shadcnBridge('.viewer', '--pptx-', {
			fallbacks: { foreground: '#f3f4f6' },
			extra: { '--office-font': 'system-ui, sans-serif' },
		});
		expect(withOptions.startsWith('.viewer {')).toBe(true);
		expect(withOptions).toContain('--office-foreground: var(--pptx-foreground, #f3f4f6);');
		expect(withOptions).toContain('--office-font: system-ui, sans-serif;');
	});
});
