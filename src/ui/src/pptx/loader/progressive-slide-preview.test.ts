// @vitest-environment happy-dom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { expect, it } from 'vitest';

import { loadPresentation } from '../dom/load-presentation';
import { createTranslator } from '../i18n';
import { createProgressiveSlidePreview } from './progressive-slide-preview';

const SAMPLE_DECK = resolve(__dirname, '../../../../../e2e/pptx/fixtures/sample-deck.pptx');

it.each([0, 3, 6])(
	'keeps the preview rail populated and isolated for initial slide %i',
	async (initialSlide) => {
		const root = document.createElement('div');
		const host = document.createElement('div');
		host.dataset.pptxLoadPreviewHost = '';
		document.body.append(root);
		const preview = createProgressiveSlidePreview({
			getRoot: () => root,
			translate: createTranslator(),
			isCurrent: () => true,
			initialSlide,
		});
		try {
			const loaded = await loadPresentation(Uint8Array.from(readFileSync(SAMPLE_DECK)).buffer, {
				onSlideLoaded(handler, progress) {
					if (progress.index === initialSlide) root.append(host);
					return preview.show(handler, progress);
				},
			});
			try {
				expect(host.querySelector('[data-pptx-opening-preview]')?.hasAttribute('aria-hidden')).toBe(
					true,
				);
				expect(host.querySelector('[data-element-id]')).toBeNull();
				expect(
					Array.from(
						host.querySelectorAll<HTMLElement>('[data-pptx-preview-thumbnail]'),
						(thumb) => thumb.dataset.pptxPreviewThumbnail,
					),
				).toEqual(['0', '1', '2', '3', '4', '5']);
				expect(loaded.slides).toHaveLength(7);
			} finally {
				loaded.handler.dispose();
			}
		} finally {
			preview.dispose();
			root.remove();
		}
		expect(host.childElementCount).toBe(0);
	},
);
