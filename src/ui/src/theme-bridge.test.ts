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
	// A product's `input` colour is the boundary of the shared form controls inside it.
	expect(aliases).toContain(
		'--_dve-input: var(--dve-input, var(--office-control-border, #8a8886));',
	);
	expect(chrome).toContain('--office-control-border: var(--_dve-input);');
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

	it('declares only the mapped tokens; derived ones resolve where they are read', () => {
		// Re-declaring `--office-select-background` here would freeze it to this selector, so a
		// host's own `--office-select-background` on a container above would stop reaching it.
		const css = themeBridge(':host > *', { '--office-background': 'var(--surface)' });
		expect(css).toBe(':host > * {\n\t--office-background: var(--surface);\n}\n');
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

	it('takes the accent ink from the product ink, else its primary', () => {
		expect(css).toContain('--office-accent-ink: var(--dve-primary-ink, var(--dve-primary));');
		// The compatibility layer prefers a host's own primary over the inherited ink, so a custom
		// primary keeps colouring text, and otherwise inherits the page's ink.
		expect(legacyThemeAliases(':host', '--xve-', '--_x-')).toContain(
			'--_x-primary-ink: var(--xve-primary-ink, var(--xve-primary, var(--office-accent-ink, light-dark(',
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
