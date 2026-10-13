import { expect, it } from 'vitest';
import type { TextSegment } from '../../core/types';
import { isContentSegment } from './text-content-style';
import { reconcileDecorationFlags } from './text-style-scope';

it('distinguishes display markers from authored numbering and excludes empty/break runs', () => {
	const literal: TextSegment = {
		text: '1. ',
		style: {},
		bulletInfo: { autoNumType: 'arabicPeriod' },
	};
	expect(isContentSegment(literal)).toBe(true);
	expect(
		isContentSegment({ ...literal, bulletInfo: { ...literal.bulletInfo, paragraphIndex: 0 } }),
	).toBe(false);
	expect(isContentSegment({ text: '• ', style: {}, bulletInfo: { char: '•' } })).toBe(false);
	for (const segment of [
		{ text: '', style: {} },
		{ text: '\n', style: {} },
		{ text: 'break', style: {}, isParagraphBreak: true },
	]) {
		expect(isContentSegment(segment)).toBe(false);
	}
});

it('reconciles inherited decorations without editing markers, breaks or source objects', () => {
	const marker: TextSegment = {
		text: '• ',
		style: { underline: false },
		bulletInfo: { char: '•' },
	};
	const breakRun: TextSegment = {
		text: '\n',
		style: {},
		isParagraphBreak: true,
		paragraphInsertionStyle: { underline: true },
	};
	const segments = [
		marker,
		{ text: 'Inherited', style: { italic: true } },
		breakRun,
		{ text: 'Off', style: { underline: false } },
	];
	const body = { underline: true, fontFamily: 'Arial' };
	const before = structuredClone({ segments, body });
	const result = reconcileDecorationFlags(segments, body, ['underline']);
	expect({ segments, body }).toEqual(before);
	expect(result.textStyle).toEqual({ fontFamily: 'Arial' });
	expect(result.textSegments[1]?.style).toEqual({ italic: true, underline: true });
	expect(result.textSegments[3]?.style.underline).toBe(false);
	expect(result.textSegments[0]).toBe(marker);
	expect(result.textSegments[2]).toBe(breakRun);
});
