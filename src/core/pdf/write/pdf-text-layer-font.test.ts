import { describe, expect, it } from 'vitest';

import { PdfFontSet } from './pdf-embedded-font';
import { PdfTextLayerFont } from './pdf-text-layer-font';
import { buildVectorPdf, PdfVectorPage } from './pdf-vector';

describe('PdfTextLayerFont', () => {
	it('gives each distinct character the next CID and maps it back to Unicode', () => {
		const font = new PdfTextLayerFont();
		expect(font.used).toBe(false);
		expect(font.encode('工作工')).toBe('000100020001');
		expect(font.encode('A')).toBe('0003');
		expect(font.used).toBe(true);
		const parts = font.build();
		expect(parts.widths).toEqual([0, 1000, 1000, 500]);
		expect(parts.toUnicode).toContain('<0001> <5de5>');
		expect(parts.toUnicode).toContain('<0002> <4f5c>');
		expect(parts.toUnicode).toContain('<0003> <0041>');
		// The font program holds six tables and no outlines.
		expect(new DataView(parts.fontFile.buffer).getUint16(4)).toBe(6);
	});

	it('measures East Asian characters one em wide and others half an em', () => {
		expect(new PdfTextLayerFont().width('工作 ab')).toBe(3.5);
	});

	it('writes characters beyond the Basic Multilingual Plane as surrogate pairs', () => {
		const font = new PdfTextLayerFont();
		font.encode('\u{20000}');
		expect(font.build().toUnicode).toContain('<0001> <d840dc00>');
	});

	it('is one font per font set, so every page shares it', () => {
		const set = new PdfFontSet();
		expect(set.textLayer).toBe(set.textLayer);
	});

	it('is written as an embedded Type0 font with a glyphless font program', () => {
		const font = new PdfTextLayerFont();
		const page = new PdfVectorPage(100, 100);
		page.embeddedText(font, 12, 0, 50, '工作总结', { invisible: true });
		expect(page.content).toContain('3 Tr');
		expect(page.content).toContain('<0001000200030004>');
		const pdf = Buffer.from(buildVectorPdf([page])).toString('latin1');
		expect(pdf).toMatch(/\/BaseFont \/[A-Z]{6}\+TextLayer/);
		expect(pdf).toContain('/Encoding /Identity-H');
		expect(pdf).toContain('/FontFile2');
		expect(pdf).toContain('/ToUnicode');
	});
});
