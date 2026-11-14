import { describeEmbeddedFontRestrictions } from 'ooxml-ui/pptx';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import FontEmbeddingPanel from './FontEmbeddingPanel.svelte';

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

describe('font embedding panel', () => {
	it('lists used and embedded fonts and updates the toggle', async () => {
		const target = document.createElement('div');
		const ontoggle = vi.fn();
		const instance = mount(FontEmbeddingPanel, {
			target,
			props: {
				usedFontFamilies: ['Aptos', 'Brand Font'],
				embeddedFonts: ['Brand Font'],
				enabled: false,
				ontoggle,
				onclose: vi.fn(),
			},
		});
		cleanup = () => unmount(instance);
		await Promise.resolve();
		flushSync();

		expect(target.textContent).toContain('Aptos');
		expect(target.textContent).toContain('Brand Font');
		const toggle = target.querySelector<HTMLInputElement>('pptx-ui-checkbox')!;
		toggle.checked = true;
		toggle.dispatchEvent(new Event('change', { bubbles: true }));
		expect(ontoggle).toHaveBeenCalledWith(true);
	});

	it('shows the licence restriction of an embedded typeface with its reason', async () => {
		const fonts = [
			{ name: 'Brand', fsType: 0x0002 },
			{ name: 'Proof', fsType: 0x0004 },
		];
		const target = document.createElement('div');
		const instance = mount(FontEmbeddingPanel, {
			target,
			props: {
				usedFontFamilies: ['Brand', 'Proof', 'Aptos'],
				embeddedFonts: fonts.map((font) => font.name),
				embeddedFontRestrictions: describeEmbeddedFontRestrictions(fonts),
				enabled: true,
				ontoggle: vi.fn(),
				onclose: vi.fn(),
			},
		});
		cleanup = () => unmount(instance);
		await vi.waitFor(() => {
			flushSync();
			expect(target.querySelector('[data-font-restriction="restricted"]')).not.toBeNull();
		});
		const restricted = target.querySelector('[data-font-restriction="restricted"]')!;
		expect(restricted.textContent).toBe('Restricted');
		expect(restricted.getAttribute('title')).toContain('does not allow it to be embedded');
		expect(
			target.querySelector('[data-font-restriction-reason="restricted"]')!.textContent,
		).toContain('a substitute is shown');
		expect(target.querySelector('[data-font-restriction="preview-print"]')!.textContent).toBe(
			'Preview and print only',
		);
		expect(target.querySelectorAll('[data-font-restriction]')).toHaveLength(2);
	});

	it('makes the toggle inert, with a reason, for a deck that embeds nothing', async () => {
		const target = document.createElement('div');
		const instance = mount(FontEmbeddingPanel, {
			target,
			props: {
				usedFontFamilies: ['Aptos'],
				embeddedFonts: [],
				enabled: false,
				canEmbed: false,
				unavailableKey: 'pptx.fonts.embedUnavailable',
				ontoggle: vi.fn(),
				onclose: vi.fn(),
			},
		});
		cleanup = () => unmount(instance);
		await Promise.resolve();
		flushSync();

		expect(target.querySelector<HTMLInputElement>('pptx-ui-checkbox')!.disabled).toBeTruthy();
		// The dictionary text, not the key: a missing entry would render the key.
		expect(target.textContent).toContain('there is nothing to embed');
	});
});
