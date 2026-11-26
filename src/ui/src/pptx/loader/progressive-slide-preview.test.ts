// @vitest-environment happy-dom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, it } from 'vitest';

import { loadPresentation } from '../dom/load-presentation';
import { createTranslator } from '../i18n';
import { createProgressiveSlidePreview } from './progressive-slide-preview';

const SAMPLE_DECK = resolve(__dirname, '../../../../../e2e/pptx/fixtures/sample-deck.pptx');

it('keeps previews isolated from the editable deck and releases the preview nodes', async () => {
	const root = document.createElement('div');
	const host = document.createElement('div');
	host.dataset.pptxLoadPreviewHost = '';
	root.append(host);
	document.body.append(root);
	const preview = createProgressiveSlidePreview({
		getRoot: () => root,
		translate: createTranslator(),
		isCurrent: () => true,
		initialSlide: 0,
	});
	try {
		const loaded = await loadPresentation(Uint8Array.from(readFileSync(SAMPLE_DECK)).buffer, {
			onSlideLoaded: preview.show,
		});
		try {
			expect(host.querySelector('[data-pptx-opening-preview]')?.hasAttribute('aria-hidden')).toBe(
				true,
			);
			expect(host.querySelector('[data-element-id]')).toBeNull();
			expect(host.querySelectorAll('[data-pptx-preview-thumbnail]')).toHaveLength(6);
			expect(loaded.slides).toHaveLength(7);
		} finally {
			loaded.handler.dispose();
		}
	} finally {
		preview.dispose();
		root.remove();
	}
	expect(host.childElementCount).toBe(0);
});
