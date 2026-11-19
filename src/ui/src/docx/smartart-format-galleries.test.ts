import { afterEach, describe, expect, it, vi } from 'vitest';

// Stands in for the 388 KB definitions entry, so each test can see whether it was requested.
const requested = vi.hoisted(() => ({ count: 0 }));
vi.mock('ooxml-core/diagram/builtin', () => {
	requested.count++;
	return {
		builtinSmartArtFormatChange: (kind: string, id: string) => ({ kind, id }),
	};
});

import { smartArtFormatGallery, syncSmartArtFormatGalleries } from './smartart-format-galleries';

const COLORS = 'urn:microsoft.com/office/officeart/2005/8/colors/accent0_1';

describe('SmartArt Design galleries in Word', () => {
	afterEach(() => document.body.replaceChildren());

	it('loads the built-in definitions only once a SmartArt can be formatted', async () => {
		const root = document.createElement('div');
		root.append(smartArtFormatGallery('colors'), smartArtFormatGallery('style'));
		document.body.append(root);
		// Building the ribbon and syncing it with nothing selected (a new document) loads nothing.
		syncSmartArtFormatGalleries(root, true);
		syncSmartArtFormatGalleries(root);
		await vi.dynamicImportSettled();
		expect(requested.count).toBe(0);

		const actions: unknown[] = [];
		root.addEventListener('ribbon-action', (event) => actions.push((event as CustomEvent).detail));
		syncSmartArtFormatGalleries(root, false);
		root
			.querySelector('[data-smartart-format="colors"]')!
			.dispatchEvent(new CustomEvent('office-gallery-pick', { detail: { itemId: COLORS } }));
		await vi.waitFor(() => expect(actions).toHaveLength(1));
		expect(requested.count).toBe(1);
		expect(actions[0]).toEqual({
			type: 'smartArtFormat',
			change: { kind: 'colors', id: COLORS },
		});
	});
});
