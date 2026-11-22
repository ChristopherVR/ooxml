import { describe, expect, it } from 'vitest';
import { withSymbolFallback } from '../../fonts/symbol-fallback';

import type { ParagraphBulletResult } from './bullet-list';
import type { ParagraphRun } from './paragraph-types';
import {
	canvasSymbolGlyphProbe,
	displayableSymbolBullet,
	displayableSymbolRuns,
} from './symbol-font-fallback';

const square: ParagraphBulletResult = {
	marker: '§',
	isNumbered: false,
	fontFamily: 'Wingdings',
	color: '#C00000',
	sizePercent: 100,
};

describe('displayableSymbolBullet (issue #290)', () => {
	it('shows a Wingdings square as ▪ in the paragraph font when the font cannot paint it', () => {
		expect(displayableSymbolBullet(square, () => false)).toEqual({
			marker: '▪',
			isNumbered: false,
			color: '#C00000',
			sizePercent: 100,
			fontFamily: withSymbolFallback('sans-serif'),
		});
	});

	it('keeps the symbol font where it paints the glyph', () => {
		expect(displayableSymbolBullet(square, () => true)).toBe(square);
	});

	it('preserves an already Unicode marker when the symbol font cannot paint it', () => {
		const arrow = { ...square, marker: '\u25ba' };
		expect(displayableSymbolBullet(arrow, () => false)).toEqual({
			...arrow,
			fontFamily: withSymbolFallback('Wingdings'),
		});
	});

	it('falls back to a plain bullet for a glyph with no known equivalent', () => {
		const unknown = { ...square, marker: '\uF07F' };
		expect(displayableSymbolBullet(unknown, () => false)?.marker).toBe('•');
	});

	it('leaves ordinary fonts, numbers and pictures alone', () => {
		const georgia = { ...square, fontFamily: 'Georgia' };
		expect(displayableSymbolBullet(georgia, () => false)).toBe(georgia);
		const numbered = { ...square, isNumbered: true };
		expect(displayableSymbolBullet(numbered, () => false)).toBe(numbered);
		expect(displayableSymbolBullet(undefined, () => false)).toBeUndefined();
	});

	it('treats the font as missing without a canvas', () => {
		expect(canvasSymbolGlyphProbe('Wingdings', '§')).toBe(false);
	});
});

describe('displayableSymbolRuns', () => {
	const run = (text: string, fontFamily?: string): ParagraphRun => ({
		text,
		style: fontFamily ? { fontFamily, color: '#111111' } : {},
	});

	it('rewrites a Wingdings checkmark run in the body font, keeping its length', () => {
		const [shown] = displayableSymbolRuns(
			[run('üü', '"Wingdings", sans-serif')],
			'Georgia',
			() => false,
		);
		expect(shown).toEqual({
			text: '✔✔',
			style: { fontFamily: withSymbolFallback('Georgia'), color: '#111111' },
		});
		expect(shown?.text).toHaveLength(2);
	});

	it('rewrites Symbol Greek and private-use a:sym glyphs inside script pieces', () => {
		const source: ParagraphRun = {
			...run('x a'),
			scriptRuns: [{ text: 'x ' }, { text: 'a', style: { fontFamily: 'Symbol' } }],
		};
		const [shown] = displayableSymbolRuns([source], 'Arial', () => false);
		expect(shown?.scriptRuns).toEqual([
			{ text: 'x ' },
			{ text: 'α•', style: { fontFamily: withSymbolFallback('Arial') } },
		]);
	});

	it('leaves runs alone where the font paints, or the font is ordinary', () => {
		const wingdings = run('ü', 'Wingdings');
		const georgia = run('ü', 'Georgia');
		const [painted, plain] = displayableSymbolRuns(
			[wingdings, georgia],
			'Arial',
			(_f, c) => c === 'ü',
		);
		expect(painted).toBe(wingdings);
		expect(plain).toBe(georgia);
	});

	it('keeps glyphs with no known equivalent', () => {
		expect(displayableSymbolRuns([run('Kü', 'Wingdings')], 'Arial', () => false)[0]?.text).toBe(
			'K✔',
		);
	});

	it('rewrites runs that inherit a symbol font from the text body', () => {
		const [shown] = displayableSymbolRuns([run('n')], 'Arial', () => false, '"Wingdings"');
		expect(shown).toEqual({ text: '■', style: { fontFamily: withSymbolFallback('Arial') } });
	});

	it('rewrites the pieces of a tabbed run along with the run', () => {
		const source: ParagraphRun = {
			...run('ü	ü', 'Wingdings'),
			tabLines: [
				{
					pieces: [
						{ text: 'ü', style: {} },
						{ text: 'ü', style: {} },
					],
				},
			],
		};
		const [shown] = displayableSymbolRuns([source], 'Arial', () => false);
		expect(shown?.tabLines?.[0]?.pieces.map((piece) => piece.text)).toEqual(['✔', '✔']);
	});
});
