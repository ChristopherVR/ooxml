import { mount } from '@vue/test-utils';
import { describeEmbeddedFontRestrictions } from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import FontEmbeddingPanel from './FontEmbeddingPanel.vue';

describe('fontEmbeddingPanel licence restrictions', () => {
	it('shows the restriction of an embedded typeface with its reason', async () => {
		const fonts = [
			{ name: 'Brand', fsType: 0x0002 },
			{ name: 'Proof', fsType: 0x0004 },
		];
		const wrapper = mount(FontEmbeddingPanel, {
			props: {
				open: true,
				embedFontsEnabled: true,
				usedFontFamilies: ['Brand', 'Proof', 'Aptos'],
				embeddedFonts: fonts.map((font) => font.name),
				embeddedFontRestrictions: describeEmbeddedFontRestrictions(fonts),
			},
			global: { stubs: { teleport: true } },
		});
		try {
			await vi.waitFor(() =>
				expect(wrapper.find('[data-font-restriction="restricted"]').exists()).toBe(true),
			);
			const restricted = wrapper.find('[data-font-restriction="restricted"]');
			expect(restricted.text()).toBe('Restricted');
			expect(restricted.attributes('title')).toContain('does not allow it to be embedded');
			expect(wrapper.find('[data-font-restriction-reason="restricted"]').text()).toContain(
				'a substitute is shown',
			);
			expect(wrapper.find('[data-font-restriction="preview-print"]').text()).toBe(
				'Preview and print only',
			);
			expect(wrapper.findAll('[data-font-restriction]')).toHaveLength(2);
		} finally {
			wrapper.unmount();
		}
	});
});
