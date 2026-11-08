/**
 * Text body options (insets, autofit, wrap), character spacing and rich
 * segments on new text boxes and shapes. Each case is saved and loaded back,
 * so the written `a:bodyPr` and runs and the reloaded model agree.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { TextSegment, TextStyle } from '../../types/text';
import { createShapeElement } from './ElementFactory';
import type { SlideBuilder } from './SlideBuilder';

interface Saved {
	body: string;
	runs: string[];
	style: TextStyle;
	segments: TextSegment[];
}

async function roundTrip(build: (slide: SlideBuilder) => void): Promise<Saved> {
	const { handler, createSlide } = await PptxHandler.create();
	const slide = createSlide('Blank');
	build(slide);
	const bytes = await handler.save([slide.build()]);
	const zip = await JSZip.loadAsync(bytes);
	const xml = (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? '';
	const loaded = await new PptxHandler().load(bytes);
	const element = loaded.slides[0]?.elements[0] as
		| { textStyle?: TextStyle; textSegments?: TextSegment[] }
		| undefined;
	return {
		body: /<a:bodyPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:bodyPr>)/u.exec(xml)?.[0] ?? '',
		runs: xml.match(/<a:r>[\s\S]*?<\/a:r>/gu) ?? [],
		style: element?.textStyle ?? {},
		segments: (element?.textSegments ?? []).filter((s) => !s.isParagraphBreak),
	};
}

describe('text body options', () => {
	it('leaves the body as before when no option is given', async () => {
		const text = await roundTrip((s) => s.addText('Hi'));
		expect(text.body).toBe('<a:bodyPr wrap="square"></a:bodyPr>');
		expect(text.style.bodyInsetLeft).toBeUndefined();
		const shape = await roundTrip((s) => s.addShape('rect', { text: 'A' }));
		expect(shape.body).not.toMatch(/Ins=|Autofit|AutoFit/u);
	});

	it('writes insets, wrap and shrink-text autofit on a text box', async () => {
		const { body, style } = await roundTrip((s) =>
			s.addText('Hi', {
				insets: { left: 0, top: 4, right: 0, bottom: 8 },
				autoFit: 'shrinkText',
				wrap: false,
			}),
		);
		expect(body).toContain('lIns="0"');
		expect(body).toContain('tIns="38100"');
		expect(body).toContain('rIns="0"');
		expect(body).toContain('bIns="76200"');
		expect(body).toContain('wrap="none"');
		expect(body).toContain('<a:normAutofit');
		expect(body).not.toContain('spAutoFit');
		expect(style.bodyInsetLeft).toBe(0);
		expect(style.bodyInsetTop).toBe(4);
		expect(style.bodyInsetBottom).toBe(8);
		expect(style.textWrap).toBe('none');
		expect(style.autoFitMode).toBe('normal');
	});

	it('writes only the inset sides that are given', async () => {
		const { body, style } = await roundTrip((s) => s.addText('Hi', { insets: { top: 2.5 } }));
		expect(body).toContain('tIns="23813"');
		expect(body).not.toMatch(/[lrb]Ins=/u);
		expect(style.bodyInsetTop).toBeCloseTo(2.5, 3);
		expect(style.bodyInsetLeft).toBeUndefined();
	});

	it('maps resizeShape to a:spAutoFit and none to a:noAutofit on shapes', async () => {
		const resize = await roundTrip((s) =>
			s.addShape('rect', { text: 'A', autoFit: 'resizeShape' }),
		);
		expect(resize.body).toContain('<a:spAutoFit');
		expect(resize.style.autoFitMode).toBe('shrink');
		const none = await roundTrip((s) => s.addShape('rect', { text: 'A', autoFit: 'none' }));
		expect(none.body).toContain('<a:noAutofit');
		expect(none.style.autoFit).toBeFalsy();
	});

	it('keeps body options on a shape without text', async () => {
		const { body, style } = await roundTrip((s) =>
			s.addShape('rect', { insets: { left: 10 }, wrap: false }),
		);
		expect(body).toContain('lIns="95250"');
		expect(body).toContain('wrap="none"');
		expect(style.bodyInsetLeft).toBe(10);
	});

	it('rejects an inset or spacing that cannot be written', async () => {
		const { createSlide } = await PptxHandler.create();
		const slide = createSlide('Blank');
		expect(() => slide.addText('x', { insets: { left: Number.NaN } })).toThrow(RangeError);
		expect(() => slide.addShape('rect', { insets: { top: 1e9 } })).toThrow(RangeError);
		expect(() =>
			slide.addText([{ text: 'x', style: { characterSpacing: Number.POSITIVE_INFINITY } }]),
		).toThrow(RangeError);
		expect(() => slide.addText('x', { autoFit: 'grow' as unknown as 'none' })).toThrow(RangeError);
	});
});

describe('character spacing', () => {
	it('writes px as hundredths of a point, rounded, and reads it back', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addText([
				{ text: 'A', style: { characterSpacing: 2 } },
				{ text: 'B', style: { characterSpacing: 0.5 } },
				{ text: 'C', style: { characterSpacing: -1.5 } },
			]),
		);
		expect(runs[0]).toContain('spc="150"');
		expect(runs[1]).toContain('spc="38"');
		expect(runs[2]).toContain('spc="-112"');
		expect(segments.map((s) => s.style?.characterSpacing)).toEqual([150, 38, -112]);
	});

	it('writes no spc for zero spacing', async () => {
		const { runs } = await roundTrip((s) =>
			s.addShape('rect', { text: 'A', textStyle: { characterSpacing: 0 } }),
		);
		expect(runs[0]).not.toContain('spc=');
	});
});

describe('rich text in shapes', () => {
	it('keeps the element text a plain string', () => {
		const el = createShapeElement('rect', {
			text: [{ text: 'A', style: { bold: true } }, { text: 'B' }],
		});
		expect(el.text).toBe('AB');
		expect(el.textSegments?.map((s) => s.text)).toEqual(['A', 'B']);
	});

	it('writes every segment with its own style on top of textStyle', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addShape('rect', {
				fill: { type: 'none' },
				text: [
					{ text: 'A', style: { bold: true } },
					{ text: 'B', style: { strikethrough: true, color: '#FF0000' } },
				],
				textStyle: { fontSize: 20 },
			}),
		);
		const [a, b] = runs;
		expect(a).toContain('b="1"');
		expect(a).not.toContain('strike="sngStrike"');
		expect(b).toContain('strike="sngStrike"');
		expect(b).toContain('<a:srgbClr val="FF0000"');
		expect(b).not.toContain('b="1"');
		expect(a).toContain('sz="1500"');
		expect(b).toContain('sz="1500"');
		expect(segments.map((s) => s.text)).toEqual(['A', 'B']);
	});

	it('keeps a single bold segment bold in a shape', async () => {
		const { runs, segments } = await roundTrip((s) =>
			s.addShape('rect', { text: [{ text: 'Bold', style: { bold: true } }] }),
		);
		expect(runs).toHaveLength(1);
		expect(runs[0]).toContain('b="1"');
		expect(segments[0]?.style?.bold).toBe(true);
	});

	it('treats an empty segment list as no text', async () => {
		const { runs, segments } = await roundTrip((s) => s.addShape('rect', { text: [] }));
		expect(runs.every((run) => run.includes('<a:t></a:t>'))).toBe(true);
		expect(segments.every((s) => s.text === '')).toBe(true);
	});

	it('splits paragraphs at a newline inside a segment', async () => {
		const { runs } = await roundTrip((s) =>
			s.addShape('rect', {
				text: [
					{ text: 'one\ntwo', style: { bold: true } },
					{ text: ' three', style: { italic: true } },
				],
			}),
		);
		expect(runs).toHaveLength(3);
	});
});
