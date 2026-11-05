// @vitest-environment happy-dom
import type { PptxEmbeddedFont } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import { createEmbeddedFontsSheet } from './embedded-fonts-sheet';

const font = (name: string): PptxEmbeddedFont =>
	({
		name,
		dataUrl: 'data:font/ttf;base64,AAEAAAAA',
		format: 'truetype',
		bold: false,
		italic: false,
	}) as PptxEmbeddedFont;

afterEach(() => {
	document.head.replaceChildren();
});

describe('createEmbeddedFontsSheet', () => {
	it('registers the deck fonts as @font-face rules in one managed style element', () => {
		const sheet = createEmbeddedFontsSheet(document, 'test-embedded-fonts');
		sheet.set([font('Zzz Face')]);
		const style = document.getElementById('test-embedded-fonts');
		expect(style?.textContent).toContain('font-family: "Zzz Face"');
		sheet.set([font('Other Face')]);
		expect(document.querySelectorAll('#test-embedded-fonts')).toHaveLength(1);
		expect(document.getElementById('test-embedded-fonts')?.textContent).toContain('Other Face');
	});

	it('removes the element for an empty list and on dispose', () => {
		const sheet = createEmbeddedFontsSheet(document, 'test-embedded-fonts');
		sheet.set([font('Zzz Face')]);
		sheet.set([]);
		expect(document.getElementById('test-embedded-fonts')).toBeNull();
		sheet.set([font('Zzz Face')]);
		sheet.dispose();
		expect(document.getElementById('test-embedded-fonts')).toBeNull();
	});

	it('exposes the rule text and families, and appends host CSS after the deck rules', () => {
		const sheet = createEmbeddedFontsSheet(document, 'test-embedded-fonts');
		sheet.set([font('Zzz Face')]);
		expect(sheet.fontFaceCss).toContain('font-family: "Zzz Face"');
		expect(sheet.fontFamilies).toHaveLength(1);
		sheet.setExtraCss('@font-face { font-family: "Host"; }');
		const css = document.getElementById('test-embedded-fonts')?.textContent ?? '';
		expect(css.indexOf('Zzz Face')).toBeLessThan(css.indexOf('Host'));
		// Host CSS alone keeps the element; clearing both removes it.
		sheet.set([]);
		expect(document.getElementById('test-embedded-fonts')?.textContent).toContain('Host');
		sheet.setExtraCss('');
		expect(document.getElementById('test-embedded-fonts')).toBeNull();
		sheet.dispose();
	});
});
