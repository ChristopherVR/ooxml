import { describe, expect, it } from 'vitest';
import { createCustomizationController } from './customization-controller';
import { resolveCustomization } from './customization-resolve';
import { initialInspectorOpen } from './inspector-customization';
import { ribbonCustomizationCss } from './ribbon-control-visibility';

describe('inspector customization (#61)', () => {
	it('allows explicit initial visibility while preserving mobile defaults', () => {
		expect(initialInspectorOpen(undefined)).toBe(true);
		expect(initialInspectorOpen(undefined, false)).toBe(false);
		expect(initialInspectorOpen(false)).toBe(false);
		expect(initialInspectorOpen(true, false)).toBe(true);
	});
	it('scopes card hiding per viewer and supports live updates and reset', () => {
		const controller = createCustomizationController({
			inspector: { hiddenCards: ['theme', 'slideSize'] },
		});
		const css = () => ribbonCustomizationCss(controller.getResolved(), 'viewer1');
		expect(css()).toContain(
			'[data-pptx-ribbon-scope="viewer1"] [data-pptx-inspector-card="theme"]',
		);
		expect(css()).toContain('[data-pptx-inspector-card="slideSize"]');
		controller.api.updateCustomization({ inspector: { hiddenCards: ['presentation'] } });
		expect(css()).not.toContain('"theme"');
		expect(css()).toContain('"presentation"');
		controller.api.setCustomization({});
		expect(css()).toBe('');
	});
	it('hides and reorders Format pane sections', () => {
		const css = ribbonCustomizationCss(
			resolveCustomization({
				inspector: {
					hiddenSections: ['effects', 'bogus"]{}' as never],
					sectionOrder: ['sizePosition', 'arrange'],
				},
			}),
			'v1',
		);
		expect(css).toContain('[data-pptx-ribbon-scope="v1"] [data-pptx-format-section="effects"]');
		expect(css).not.toContain('bogus');
		expect(css).toContain('[data-pptx-format-section="sizePosition"] {\n\torder: -2;');
		expect(css).toContain('[data-pptx-format-section="arrange"] {\n\torder: -1;');
	});
	it('ignores unrecognized card ids from untyped hosts', () => {
		expect(
			ribbonCustomizationCss(
				resolveCustomization({ inspector: { hiddenCards: ['theme"]{} body{' as never] } }),
			),
		).toBe('');
	});
});
