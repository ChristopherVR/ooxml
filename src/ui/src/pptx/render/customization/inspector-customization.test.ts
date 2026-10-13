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
	it('ignores unrecognized card ids from untyped hosts', () => {
		expect(
			ribbonCustomizationCss(
				resolveCustomization({ inspector: { hiddenCards: ['theme"]{} body{' as never] } }),
			),
		).toBe('');
	});
});
