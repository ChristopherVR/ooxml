import { describe, expect, it } from 'vitest';
import { withOfficeThemeCss, pptxThemeValue } from './office-css';
describe('Office theme compatibility', () => {
	it('resolves inherited legacy overrides at the current chrome scope without declaring raw defaults', () => {
		const css = withOfficeThemeCss(
			'.pptxv { color:var(--pptx-foreground); border:1px solid var(--pptx-border,#333); }',
			'.pptxv',
		);
		expect(css).toContain('--_pptx-foreground: var(--pptx-foreground, var(--office-foreground,');
		expect(css).toContain('color:var(--_pptx-foreground)');
		expect(css).toContain('var(--_pptx-border, #333)');
		expect(css).not.toMatch(/(?:^|[;{]\s*)--pptx-background\s*:/);
		expect(withOfficeThemeCss(css, '.pptxv')).toBe(css);
	});
	it('preserves public legacy names while giving static utility CSS canonical fallbacks', () => {
		expect(pptxThemeValue('background')).toContain(
			'var(--pptx-background, var(--office-background,',
		);
		expect(pptxThemeValue('primary')).toContain('var(--office-accent,');
	});
});

it('keeps partial legacy fallback roles ahead of canonical defaults', () => {
	const css = withOfficeThemeCss(
		'a{background:var(--pptx-popover,var(--pptx-card,#fff));}',
		':host',
	);
	expect(css).toContain(
		'background:var(--pptx-popover, var(--pptx-card, var(--_pptx-popover, #fff)))',
	);
});
