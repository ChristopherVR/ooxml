import { describe, expect, it } from 'vitest';
import type { TextSegment, TextStyle } from '../core';
import { drawTextSegments } from './ShapeTextRenderer';

function paint(segments: TextSegment[], style: TextStyle = {}, scale = 1) {
	const calls: { text: string; x: number; y: number; font: string; color: string }[] = [];
	const context = {
		font: '',
		fillStyle: '',
		measureText: (text: string) => ({ width: text.length * 10 }),
		fillText(text: string, x: number, y: number) {
			calls.push({ text, x, y, font: this.font, color: this.fillStyle });
		},
	};
	drawTextSegments(
		context as unknown as CanvasRenderingContext2D,
		segments,
		style,
		200,
		200,
		scale,
	);
	return calls;
}

const run = (text: string): TextSegment => ({ text, style: { fontSize: 10 } });
const zeroInsets: TextStyle = {
	bodyInsetLeft: 0,
	bodyInsetRight: 0,
	bodyInsetTop: 0,
	bodyInsetBottom: 0,
};

describe('shape text raster layout', () => {
	it('uses body and break font sizes for consecutive empty paragraphs', () => {
		const calls = paint(
			[
				{ text: 'A', style: {} },
				{ text: '', style: { fontSize: 24 }, isParagraphBreak: true },
				{ text: '', style: {}, isParagraphBreak: true },
				{ text: 'B', style: {} },
			],
			{ ...zeroInsets, fontSize: 48 },
		);
		expect(calls[0]!.y).toBeCloseTo(57.6);
		expect(calls[1]!.y - calls[0]!.y).toBeCloseTo((57.6 + 28.8) * 1.3);
	});

	it('inherits body properties when a run contains explicit undefined values', () => {
		const calls = paint([{ text: 'A', style: { fontSize: undefined, color: undefined } }], {
			...zeroInsets,
			fontSize: 48,
			color: '#123456',
		});
		expect(calls[0]).toMatchObject({ font: 'normal normal 48px sans-serif', color: '#123456' });
	});

	it('honors explicit zero insets and scales nonzero insets', () => {
		expect(paint([run('A')], zeroInsets)[0]).toMatchObject({ x: 0, y: 12 });
		expect(
			paint([run('A')], { ...zeroInsets, bodyInsetLeft: 5, bodyInsetTop: 7 }, 2)[0],
		).toMatchObject({ x: 10, y: 38 });
	});

	it('keeps soft breaks as line breaks instead of spaces', () => {
		const calls = paint(
			[run('A'), { text: '', style: {}, isLineBreak: true }, run('B')],
			zeroInsets,
		);
		expect(calls.map((call) => call.text)).toEqual(['A', 'B']);
		expect(calls[1]!.y).toBeGreaterThan(calls[0]!.y);
		expect(calls[1]!.x).toBe(0);
	});

	it('handles newline text, paragraph separators and consecutive empty paragraphs', () => {
		const calls = paint(
			[
				run('A\r\nB'),
				{ text: '\n', style: {} },
				{ text: '', style: {}, isParagraphBreak: true },
				run('C'),
			],
			zeroInsets,
		);
		expect(calls.map((call) => call.text)).toEqual(['A', 'B', 'C']);
		expect(calls[2]!.y - calls[1]!.y).toBeGreaterThan(calls[1]!.y - calls[0]!.y);
	});

	it('inherits body font and color while preserving explicit run overrides', () => {
		const calls = paint(
			[
				{ text: 'A', style: {} },
				{ text: 'B', style: { fontSize: 10, color: '#ff0000', bold: false } },
			],
			{ ...zeroInsets, fontSize: 24, bold: true, color: '#0000ff' },
		);
		expect(calls[0]).toMatchObject({ font: 'normal bold 24px sans-serif', color: '#0000ff' });
		expect(calls[1]).toMatchObject({ font: 'normal normal 10px sans-serif', color: '#ff0000' });
	});
});
