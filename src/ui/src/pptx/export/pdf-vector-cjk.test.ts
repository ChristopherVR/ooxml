// @vitest-environment happy-dom
/**
 * Chinese, Japanese and Korean text in a deck that embeds no font for it (no standard PDF font
 * has those characters): the collector asks for a picture of the run and the painter keeps the
 * text as an invisible, searchable layer, instead of leaving it out of the file.
 */
import { inflateSync } from 'node:zlib';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PdfFontSet } from 'ooxml-core/pdf/write';
import { buildVectorPdf } from 'ooxml-core/pdf/write';
import { syntheticTrueType } from '../../../../core/text/truetype/test-support/synthetic-truetype';
import { IDENTITY, type TextItem } from './pdf-vector-display-list';
import { canWritePdfText, paintSlide } from './pdf-vector-paint';
import { measureTextNode } from './pdf-vector-text';

const latin1 = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');

function streams(bytes: Uint8Array): string {
	const text = latin1(bytes);
	const out: string[] = [];
	for (const match of text.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const start = match.index! + match[0].length;
		const data = bytes.subarray(start, start + length);
		if (/\/Length1/.test(match[1]!)) continue;
		out.push(
			match[1]!.includes('/FlateDecode') ? latin1(inflateSync(Buffer.from(data))) : latin1(data),
		);
	}
	return out.join('\n');
}

const size = { canvasWidth: 1280, canvasHeight: 720, pageWidth: 960, pageHeight: 540 };
const FAMILY = '"微软雅黑", sans-serif';

const item = (overrides: Partial<TextItem> = {}): TextItem => ({
	kind: 'text',
	matrix: IDENTITY,
	clips: [],
	alpha: 1,
	text: '工作总结',
	x: 100,
	baseline: 200,
	width: 160,
	size: 40,
	family: FAMILY,
	bold: false,
	italic: false,
	color: { r: 1, g: 1, b: 1, a: 1 },
	...overrides,
});

/** A font with glyphs for 工 and 作 (and space, A, B). */
const cjkFont = () =>
	new PdfFontSet([
		{
			family: '微软雅黑',
			data: syntheticTrueType([
				[0x4f5c, 1],
				[0x5de5, 2],
			]),
		},
	]);

describe('vector PDF text no font can write', () => {
	it('knows CJK needs a font that has it, and Latin does not', () => {
		expect(canWritePdfText('CONTENTS 01', FAMILY, false, false)).toBe(true);
		expect(canWritePdfText('工作总结', FAMILY, false, false)).toBe(false);
		expect(canWritePdfText('目次', FAMILY, false, false, new PdfFontSet())).toBe(false);
		expect(canWritePdfText('목차', FAMILY, false, false)).toBe(false);
		expect(canWritePdfText('工作', FAMILY, false, false, cjkFont())).toBe(true);
		expect(canWritePdfText('工作总结', FAMILY, false, false, cjkFont())).toBe(false);
	});

	it('keeps invisible CJK text over its picture as a searchable layer', () => {
		const fonts = new PdfFontSet();
		const { page, diagnostics } = paintSlide([item({ invisible: true })], { ...size, fonts });
		expect(diagnostics).toEqual([]);
		expect(page.content).toMatch(/\/E0 40 Tf .* 3 Tr \[<0001000200030004>\] TJ ET.* 0 Tr/);
		expect([...page.embeddedFonts.values()]).toEqual([fonts.textLayer]);
		const bytes = buildVectorPdf([page]);
		const decoded = streams(bytes);
		for (const [cid, code] of [
			['0001', '5de5'],
			['0002', '4f5c'],
			['0003', '603b'],
			['0004', '7ed3'],
		])
			expect(decoded).toContain(`<${cid}> <${code}>`);
		expect(latin1(bytes)).toMatch(/\/BaseFont \/[A-Z]{6}\+TextLayer/);
	});

	it('fits the invisible layer to the measured width of the run', () => {
		// Four one-em characters at 40 px measure 160 px natural; 320 px is 200 %.
		const { page } = paintSlide([item({ invisible: true, width: 320 })], size);
		expect(page.content).toContain(' 200 Tz');
	});

	it('reports visible text no font has, and still keeps it searchable', () => {
		const { page, diagnostics } = paintSlide([item()], size);
		expect(diagnostics).toEqual([{ code: 'unsearchableText' }]);
		expect(page.content).toContain('3 Tr');
	});

	it('writes CJK as real, visible text when a supplied font has every character', () => {
		const fonts = cjkFont();
		const { page, diagnostics } = paintSlide([item({ text: '工作' })], { ...size, fonts });
		expect(diagnostics).toEqual([]);
		expect(page.content).not.toContain('3 Tr');
		expect([...page.embeddedFonts.values()]).toEqual([fonts.fonts[0]]);
	});
});

describe('measuring text no font can write', () => {
	afterEach(() => vi.restoreAllMocks());

	const style = (fontFamily: string) =>
		({
			writingMode: 'horizontal-tb',
			textTransform: 'none',
			getPropertyValue: () => '',
			color: 'rgb(255, 255, 255)',
			whiteSpace: 'normal',
			fontSize: '40px',
			fontWeight: '400',
			fontStyle: 'normal',
			fontFamily,
			textDecorationLine: 'none',
			textAlign: 'center',
		}) as unknown as CSSStyleDeclaration;

	/** A text node whose every range measures as one box (happy-dom lays nothing out). */
	function textNode(data: string): Text {
		const rect = { left: 100, right: 260, top: 170, width: 160, height: 50 };
		vi.spyOn(document, 'createRange').mockReturnValue({
			setStart: () => undefined,
			setEnd: () => undefined,
			getClientRects: () => [rect],
		} as unknown as Range);
		return document.createTextNode(data);
	}

	const context = (fonts?: PdfFontSet) => ({
		matrix: IDENTITY,
		clips: [],
		alpha: 1,
		originX: 0,
		originY: 0,
		scale: 1,
		...(fonts ? { fonts } : {}),
	});

	it('asks for a picture of CJK text with the text kept as invisible items', () => {
		const result = measureTextNode(
			textNode('工作总结'),
			style(FAMILY),
			context(new PdfFontSet()),
			100,
		);
		expect(result.raster).toBe('noFontTextAsPicture');
		expect(result.items).toHaveLength(1);
		expect(result.items[0]).toMatchObject({ text: '工作总结', invisible: true, width: 160 });
	});

	it('asks for a picture of Japanese and Korean text too', () => {
		expect(measureTextNode(textNode('目次'), style(FAMILY), context(), 100).raster).toBe(
			'noFontTextAsPicture',
		);
		expect(measureTextNode(textNode('목차'), style(FAMILY), context(), 100).raster).toBe(
			'noFontTextAsPicture',
		);
	});

	it('keeps Latin text as visible text', () => {
		const result = measureTextNode(textNode('CONTENTS'), style('Calibri'), context(), 100);
		expect(result.raster).toBeUndefined();
		expect(result.items[0]?.invisible).toBeUndefined();
	});

	it('keeps CJK as visible text when a supplied font covers it', () => {
		const result = measureTextNode(textNode('工作'), style(FAMILY), context(cjkFont()), 100);
		expect(result.raster).toBeUndefined();
		expect(result.items[0]?.invisible).toBeUndefined();
	});

	it('keeps vertical text as one invisible item per character over its picture', () => {
		const vertical = { ...style(FAMILY), writingMode: 'vertical-rl' } as CSSStyleDeclaration;
		const result = measureTextNode(textNode('目录'), vertical, context(), 100);
		expect(result.raster).toBe('verticalTextAsPicture');
		expect(result.items.map((entry) => [entry.text, entry.invisible])).toEqual([
			['目', true],
			['录', true],
		]);
	});

	it('asks for no second picture inside one already drawn as a picture', () => {
		const result = measureTextNode(
			textNode('工作总结'),
			style(FAMILY),
			{ ...context(), invisible: true },
			100,
		);
		expect(result.raster).toBeUndefined();
		expect(result.items[0]?.invisible).toBe(true);
	});
});
