/**
 * text-style-scope.test.ts: a `TextStyle` update carries both body-level
 * (`a:bodyPr` / paragraph geometry the block builder reads from
 * `element.textStyle` only) and run-level keys. When an inline selection is
 * active the run keys go to the selected runs and the body keys MUST still
 * reach `element.textStyle`, or Increase Indent does nothing.
 */

import type { PptxElement, TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	BODY_TEXT_STYLE_KEYS,
	isBodyTextStyleKey,
	overrideParagraphIndents,
	overrideParagraphProperties,
	paragraphOverridePatch,
	reconcileDecorationFlags,
	splitTextStyleUpdate,
	withExclusiveLineSpacing,
} from './text-style-scope';

const seg = (text: string, style: TextSegment['style'] = {}): TextSegment => ({ text, style });

describe('body text style keys', () => {
	it('names every key text-block-style reads from element.textStyle only', () => {
		for (const key of [
			'paragraphMarginLeft',
			'paragraphMarginRight',
			'paragraphIndent',
			'vAlign',
			'textDirection',
			'columnCount',
			'columnSpacing',
			'textWrap',
			'autoFit',
			'autoFitMode',
			'bodyInsetLeft',
			'bodyInsetTop',
			'bodyInsetRight',
			'bodyInsetBottom',
			'lineSpacing',
			'paragraphSpacingBefore',
			'paragraphSpacingAfter',
			'textWarpPreset',
		]) {
			expect(isBodyTextStyleKey(key)).toBeTruthy();
		}
	});

	it('leaves run-level keys to the selected runs', () => {
		for (const key of ['bold', 'italic', 'underline', 'strikethrough', 'fontSize', 'color']) {
			expect(BODY_TEXT_STYLE_KEYS.has(key as never)).toBeFalsy();
		}
	});
});

describe('splitTextStyleUpdate', () => {
	it('partitions an update into body and run halves', () => {
		const split = splitTextStyleUpdate({
			paragraphMarginLeft: 24,
			bold: true,
			vAlign: 'middle',
			fontSize: 12,
		});
		expect(split.body).toStrictEqual({ paragraphMarginLeft: 24, vAlign: 'middle' });
		expect(split.run).toStrictEqual({ bold: true, fontSize: 12 });
	});

	it('keeps an explicit undefined so a clear still reaches its target', () => {
		const split = splitTextStyleUpdate({ highlightColor: undefined, textWrap: undefined });
		expect('highlightColor' in split.run).toBeTruthy();
		expect('textWrap' in split.body).toBeTruthy();
	});
});

describe('reconcileDecorationFlags', () => {
	it('drops the body flag once a run no longer carries it, keeping the others decorated', () => {
		const segments = [seg('keep'), seg('plain', { underline: false })];
		const next = reconcileDecorationFlags(segments, { underline: true, fontSize: 18 }, [
			'underline',
		]);
		expect(next.textStyle).toStrictEqual({ fontSize: 18 });
		expect(next.textSegments[0].style.underline).toBeTruthy();
		expect(next.textSegments[1].style.underline).toBeFalsy();
	});

	it('promotes the body flag when every content run agrees', () => {
		const segments = [
			seg('a', { bold: true }),
			{ text: '\n', style: {}, isParagraphBreak: true },
			seg('b', { bold: true }),
		];
		const next = reconcileDecorationFlags(segments, {}, ['bold']);
		expect(next.textStyle.bold).toBeTruthy();
		expect(next.textSegments).toStrictEqual(segments);
	});

	it('only touches the flags it is asked about', () => {
		const segments = [seg('a', { italic: false })];
		const next = reconcileDecorationFlags(segments, { italic: true, bold: true }, ['bold']);
		expect(next.textStyle).toStrictEqual({ italic: true, bold: true });
	});
});

describe('withExclusiveLineSpacing', () => {
	it('clears the exact size when a multiplier is set, and the reverse', () => {
		expect(withExclusiveLineSpacing({ lineSpacing: 2 })).toStrictEqual({
			lineSpacing: 2,
			lineSpacingExactPt: undefined,
		});
		expect(withExclusiveLineSpacing({ lineSpacingExactPt: 12 })).toStrictEqual({
			lineSpacingExactPt: 12,
			lineSpacing: undefined,
		});
	});

	it('leaves an update that names both modes, or neither, alone', () => {
		expect(withExclusiveLineSpacing({ lineSpacing: 1, lineSpacingExactPt: 9 })).toStrictEqual({
			lineSpacing: 1,
			lineSpacingExactPt: 9,
		});
		expect(withExclusiveLineSpacing({ bold: true })).toStrictEqual({ bold: true });
	});
});

describe('overrideParagraphProperties', () => {
	const para = (paragraphProperties: TextSegment['paragraphProperties']): TextSegment => ({
		text: 'a',
		style: {},
		paragraphProperties,
	});

	it('replaces an authored exact line spacing with the new multiplier', () => {
		const [out] = overrideParagraphProperties(
			[para({ lineSpacingExactPt: 14.3, paragraphMarginLeft: 0 })],
			withExclusiveLineSpacing({ lineSpacing: 2 }),
		);
		expect(out?.paragraphProperties).toStrictEqual({ paragraphMarginLeft: 0, lineSpacing: 2 });
	});

	it('rewrites an authored indent', () => {
		const [out] = overrideParagraphProperties([para({ paragraphMarginLeft: 0 })], {
			paragraphMarginLeft: 24,
		});
		expect(out?.paragraphProperties).toStrictEqual({ paragraphMarginLeft: 24 });
	});

	it('pins nothing new onto a paragraph that does not declare the key', () => {
		const segment = para({ align: 'left' });
		const plain = seg('b');
		const out = overrideParagraphProperties([segment, plain], {
			lineSpacing: 2,
			paragraphIndent: 5,
		});
		expect(out[0]).toBe(segment);
		expect(out[1]).toBe(plain);
	});
});

describe('paragraph overrides leave run-level keys alone', () => {
	it('does not rewrite an authored alignment', () => {
		const segment: TextSegment = { text: 'a', style: {}, paragraphProperties: { align: 'left' } };
		expect(overrideParagraphProperties([segment], { align: 'center' })[0]).toBe(segment);
	});
});

describe('overrideParagraphIndents', () => {
	const el = (paragraphIndents?: { marginLeft?: number; indent?: number }[]) =>
		({ type: 'text', id: 't', x: 0, y: 0, width: 10, height: 10, paragraphIndents }) as PptxElement;

	it('applies an element-wide indent to every paragraph the renderer indents', () => {
		const out = overrideParagraphIndents(el([{ marginLeft: 0, indent: -6 }, {}]), {
			paragraphMarginLeft: 24,
		});
		expect(out.paragraphIndents).toStrictEqual([
			{ marginLeft: 24, indent: -6 },
			{ marginLeft: 24 },
		]);
	});

	it('is empty without parsed indents or an indent update', () => {
		expect(overrideParagraphIndents(el(), { paragraphMarginLeft: 24 })).toStrictEqual({});
		expect(overrideParagraphIndents(el([{}]), { lineSpacing: 2 })).toStrictEqual({});
	});
});

describe('paragraphOverridePatch', () => {
	it('sends no segments when no paragraph declares the changed key', () => {
		const el = {
			type: 'text',
			id: 't',
			x: 0,
			y: 0,
			width: 10,
			height: 10,
			textSegments: [{ text: 'a', style: {}, paragraphProperties: { align: 'left' } }],
		} as PptxElement;
		expect(paragraphOverridePatch(el, { vAlign: 'middle' })).toStrictEqual({});
	});
});
