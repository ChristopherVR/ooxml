// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n/translator';
import { PRESETS } from '../render/smart-art-presets';
import { buildSmartArtInsertPreview } from './smartart-insert-preview';

const TILE = { width: 64, height: 40 };

describe('buildSmartArtInsertPreview', () => {
	it('renders the real SmartArt output for the preset, not a generic icon', () => {
		const tile = buildSmartArtInsertPreview(document, createTranslator(), PRESETS[0]!.layout, TILE);
		expect(tile.querySelector('.pptxv-smartart')).toBeTruthy();
	});

	it('scales the full-size preview element into the requested box', () => {
		const tile = buildSmartArtInsertPreview(document, createTranslator(), PRESETS[0]!.layout, TILE);
		expect(Number.parseFloat(tile.style.width)).toBeLessThanOrEqual(64);
		expect(Number.parseFloat(tile.style.height)).toBeLessThanOrEqual(40);
		expect((tile.firstElementChild as HTMLElement).style.transform).toContain('scale(');
	});

	it('renders a distinct diagram per layout', () => {
		const t = createTranslator();
		const a = buildSmartArtInsertPreview(document, t, PRESETS[0]!.layout, TILE);
		const b = buildSmartArtInsertPreview(document, t, PRESETS[1]!.layout, TILE);
		expect(a.querySelector('.pptxv-smartart')?.getAttribute('data-testid')).not.toBe(
			b.querySelector('.pptxv-smartart')?.getAttribute('data-testid'),
		);
	});
});
