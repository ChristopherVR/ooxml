import { describe, expect, it } from 'vitest';
import { shadcnBridge, themeBridge } from './theme-bridge.js';

describe('themeBridge', () => {
	it('declares each mapped token on the selector and skips empty ones', () => {
		const css = themeBridge(':host', {
			'--office-accent': 'var(--brand)',
			'--office-border': '',
		});
		expect(css).toBe(':host {\n\t--office-accent: var(--brand);\n}\n');
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
