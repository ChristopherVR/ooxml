import type { TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	htmlToPasteSegments,
	segmentsBodyText,
	spliceTextSegments,
	styleAtOffset,
} from './html-text-paste';

const destination = { fontFamily: 'Calibri', fontSize: 18, color: '#333333', bold: false };

describe('htmlToPasteSegments', () => {
	it('keeps source formatting over the destination style by default', () => {
		const pasted = htmlToPasteSegments(
			'<p><b>Bold</b> <span style="color:#ff0000;font-family:Georgia">red</span></p>',
			destination,
		);
		expect(pasted?.rich).toBe(true);
		expect(pasted?.text).toBe('Bold red');
		expect(pasted?.segments[0]?.style).toEqual({ ...destination, bold: true });
		expect(pasted?.segments[2]?.style).toMatchObject({
			color: '#FF0000',
			fontFamily: 'Georgia',
			fontSize: 18,
		});
	});

	it('drops pasted fonts and colours for Use Destination Theme', () => {
		const pasted = htmlToPasteSegments(
			'<p><b style="color:red;font-family:Georgia">x</b></p>',
			destination,
			'use-destination-theme',
		);
		expect(pasted?.segments[0]?.style).toEqual({ ...destination, bold: true });
	});

	it('keeps only text and lines for Keep Text Only', () => {
		const pasted = htmlToPasteSegments(
			'<ul><li><b>a</b></li><li>b</li></ul>',
			destination,
			'keep-text-only',
		);
		expect(pasted?.rich).toBe(false);
		expect(pasted?.segments.map((s) => s.text)).toEqual(['a', '\n', 'b']);
		expect(pasted?.segments.every((s) => !s.bulletInfo)).toBe(true);
	});

	it('maps bullets, numbering, levels, alignment, links, highlight and baseline', () => {
		const pasted = htmlToPasteSegments(
			'<p>lead</p><ol><li style="text-align:center">one</li></ol><ul><li><ul><li><a href="https://e.x/">l</a></li></ul></li></ul>' +
				'<p><mark>h</mark>x<sup>2</sup></p>',
			destination,
		);
		const segments = pasted?.segments ?? [];
		const first = (text: string) => segments.find((s) => s.text === text);
		expect(first('one')).toMatchObject({
			bulletInfo: { autoNumType: 'arabicPeriod' },
			paragraphProperties: { align: 'center', paragraphMarginLeft: 30, paragraphIndent: -30 },
		});
		expect(first('l')).toMatchObject({
			bulletInfo: { char: '–' },
			paragraphLevel: 1,
			style: { hyperlink: 'https://e.x/' },
		});
		expect(first('h')?.style.highlightColor).toBe('#FFFF00');
		expect(first('2')?.style.baseline).toBe(30000);
		expect(first('lead')?.bulletInfo).toBeUndefined();
	});

	it('reports plain HTML as not rich so the caller pastes text natively', () => {
		expect(htmlToPasteSegments('<span>plain</span>', destination)?.rich).toBe(false);
		expect(htmlToPasteSegments('<p></p>', destination)).toBeUndefined();
	});
});

describe('spliceTextSegments', () => {
	const body: TextSegment[] = [
		{ text: 'Hello ', style: { bold: true }, bulletInfo: { char: '-' } },
		{ text: 'world', style: { italic: true } },
		{ text: '\n', style: {}, isParagraphBreak: true },
		{ text: 'Next', style: {}, paragraphLevel: 1 },
	];

	it('replaces a selection inside a run and keeps the paragraph properties', () => {
		const result = spliceTextSegments(body, 6, 9, [{ text: 'X', style: { color: '#FF0000' } }]);
		expect(segmentsBodyText(result)).toBe('Hello Xld\nNext');
		expect(result[0]?.bulletInfo).toEqual({ char: '-' });
		expect(result[2]).toEqual({ text: 'ld', style: { italic: true } });
	});

	it('moves the destination paragraph properties onto text pasted at its start', () => {
		const result = spliceTextSegments(body, 12, 12, [{ text: 'P', style: {} }]);
		expect(segmentsBodyText(result)).toBe('Hello world\nPNext');
		expect(result.find((s) => s.text === 'P')?.paragraphLevel).toBe(1);
		expect(result.find((s) => s.text === 'Next')?.paragraphLevel).toBeUndefined();
	});

	it('appends at the end of the body', () => {
		const result = spliceTextSegments(body, 99, 99, [{ text: '!', style: {} }]);
		expect(segmentsBodyText(result)).toBe('Hello world\nNext!');
	});

	it('reads the style at the caret from the run it follows', () => {
		expect(styleAtOffset(body, 3)).toEqual({ bold: true });
		expect(styleAtOffset(body, 6)).toEqual({ bold: true });
		expect(styleAtOffset(body, 7)).toEqual({ italic: true });
		expect(styleAtOffset([], 0, { fontSize: 9 })).toEqual({ fontSize: 9 });
	});
});
