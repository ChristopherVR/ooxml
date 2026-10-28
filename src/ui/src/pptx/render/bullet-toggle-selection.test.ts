/**
 * bullet-toggle-selection.test.ts: with an inline selection active, Bullets /
 * Numbering touch only the paragraphs the selection intersects (PowerPoint;
 * with `listConversionScope: 'list'`, a bare caret switching a list's kind
 * converts the whole list instead),
 * and the selection is remapped over the inserted / removed marker segments.
 */

import type { PptxElement, TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	isBulletMarkerSegment,
	paragraphBulletKind,
	splitBulletParagraphs,
	toggleElementBullets,
} from './bullet-toggle';
import {
	selectionBulletKind,
	setSelectionBullets,
	toggleSelectionBullets,
} from './bullet-toggle-selection';
import { applyListStyleUpdate } from './text-list-style-update';

const seg = (text: string, extra: Partial<TextSegment> = {}): TextSegment => ({
	text,
	style: { fontSize: 18 },
	...extra,
});
const brk = (): TextSegment => ({ text: '\n', style: {}, isParagraphBreak: true });
const loadedBullet = (text: string): TextSegment[] => [
	seg('• ', { bulletInfo: { char: '•' } }),
	seg(text),
];

const textElement = (segments: TextSegment[]): PptxElement =>
	({
		type: 'text',
		id: 't1',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		text: segments.map((s) => s.text).join(''),
		textStyle: { fontSize: 18 },
		textSegments: segments,
	}) as unknown as PptxElement;

const texts = (patch: Partial<PptxElement>): string[] =>
	(patch as { textSegments: TextSegment[] }).textSegments.map((s) => s.text);

describe('selectionBulletKind', () => {
	it('does not include the next paragraph when the range ends at its start', () => {
		const el = textElement([...loadedBullet('A'), brk(), seg('B')]);
		const selection = { startSegIdx: 1, startOffset: 0, endSegIdx: 3, endOffset: 0 };
		expect(selectionBulletKind(el, selection)).toBe('bullet');
		expect(texts(toggleSelectionBullets(el, 'bullet', selection).patch)).toStrictEqual([
			'A',
			'\n',
			'B',
		]);
	});

	it('reads the paragraphs the selection intersects, mixed when they disagree', () => {
		const el = textElement([...loadedBullet('A'), brk(), seg('B'), brk(), seg('C')]);
		expect(
			selectionBulletKind(el, { startSegIdx: 1, startOffset: 0, endSegIdx: 1, endOffset: 1 }),
		).toBe('bullet');
		expect(
			selectionBulletKind(el, { startSegIdx: 3, startOffset: 0, endSegIdx: 5, endOffset: 1 }),
		).toBe('none');
		expect(
			selectionBulletKind(el, { startSegIdx: 1, startOffset: 0, endSegIdx: 3, endOffset: 1 }),
		).toBe('mixed');
		expect(selectionBulletKind(el, null)).toBe('mixed');
	});
});

describe('toggleSelectionBullets', () => {
	it.each([
		{ index: 0, offset: 1, expectedIndex: 1, expectedOffset: 1 },
		{ index: 1, offset: 0, expectedIndex: 2, expectedOffset: 0 },
	])(
		'retains explicit run-boundary affinity at $index/$offset',
		({ index, offset, expectedIndex, expectedOffset }) => {
			const el = textElement([seg('A'), seg('B', { style: { bold: true } })]);
			const { newSelection } = setSelectionBullets(el, 'bullet', {
				startSegIdx: index,
				startOffset: offset,
				endSegIdx: index,
				endOffset: offset,
			});
			expect(newSelection).toStrictEqual({
				startSegIdx: expectedIndex,
				startOffset: expectedOffset,
				endSegIdx: expectedIndex,
				endOffset: expectedOffset,
			});
		},
	);

	it('keeps a selected empty item caret on that item after changing its marker', () => {
		const el = textElement([
			seg('A'),
			brk(),
			seg('◆ ', {
				bulletInfo: { char: '◆' },
				paragraphInsertionStyle: { fontSize: 30 },
			}),
		]);
		const { patch, newSelection } = setSelectionBullets(el, 'numbered', {
			startSegIdx: 2,
			startOffset: 0,
			endSegIdx: 2,
			endOffset: 0,
		});
		expect(texts(patch)).toStrictEqual(['A', '\n', '1.']);
		expect(newSelection).toStrictEqual({
			startSegIdx: 2,
			startOffset: 0,
			endSegIdx: 2,
			endOffset: 0,
		});
	});

	it('can apply an explicit collapsed selection at the start of a later paragraph', () => {
		const el = textElement([seg('A'), brk(), seg('B')]);
		const { patch, newSelection } = setSelectionBullets(el, 'bullet', {
			startSegIdx: 2,
			startOffset: 0,
			endSegIdx: 2,
			endOffset: 0,
		});
		expect(texts(patch)).toStrictEqual(['A', '\n', '• ', 'B']);
		expect(newSelection).toStrictEqual({
			startSegIdx: 3,
			startOffset: 0,
			endSegIdx: 3,
			endOffset: 0,
		});
	});

	it('bullets only the selected paragraph and shifts the selection past its marker', () => {
		const el = textElement([seg('A'), brk(), seg('B'), brk(), seg('C')]);
		const { patch, newSelection } = toggleSelectionBullets(el, 'bullet', {
			startSegIdx: 2,
			startOffset: 0,
			endSegIdx: 2,
			endOffset: 1,
		});
		expect(texts(patch)).toStrictEqual(['A', '\n', '• ', 'B', '\n', 'C']);
		expect(newSelection).toStrictEqual({
			startSegIdx: 3,
			startOffset: 0,
			endSegIdx: 3,
			endOffset: 1,
		});
	});

	it('applies to every paragraph a multi-paragraph selection touches, numbering consecutively', () => {
		const el = textElement([seg('A'), brk(), seg('B'), brk(), seg('C')]);
		const { patch, newSelection } = toggleSelectionBullets(el, 'numbered', {
			startSegIdx: 0,
			startOffset: 1,
			endSegIdx: 2,
			endOffset: 1,
		});
		expect(texts(patch)).toStrictEqual(['1.', 'A', '\n', '2.', 'B', '\n', 'C']);
		expect(newSelection).toStrictEqual({
			startSegIdx: 1,
			startOffset: 1,
			endSegIdx: 4,
			endOffset: 1,
		});
	});

	it('turns the selected paragraphs off when they are all already in that state', () => {
		const el = textElement([...loadedBullet('A'), brk(), ...loadedBullet('B')]);
		const { patch, newSelection } = toggleSelectionBullets(el, 'bullet', {
			startSegIdx: 4,
			startOffset: 0,
			endSegIdx: 4,
			endOffset: 1,
		});
		expect(texts(patch)).toStrictEqual(['• ', 'A', '\n', 'B']);
		expect(newSelection).toStrictEqual({
			startSegIdx: 3,
			startOffset: 0,
			endSegIdx: 3,
			endOffset: 1,
		});
		const segments = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(segments[3].bulletInfo).toStrictEqual({ char: '•', none: true });
	});

	it('numbers nested paragraphs independently at each level', () => {
		const el = textElement([
			seg('A', { paragraphLevel: 0 }),
			brk(),
			seg('B', { paragraphLevel: 1 }),
			brk(),
			seg('C', { paragraphLevel: 1 }),
			brk(),
			seg('D', { paragraphLevel: 0 }),
		]);
		const { patch } = setSelectionBullets(el, 'numbered', null);
		expect(texts(patch)).toStrictEqual([
			'1.',
			'A',
			'\n',
			'1.',
			'B',
			'\n',
			'2.',
			'C',
			'\n',
			'2.',
			'D',
		]);
	});

	it('converts an empty listed paragraph and retains its insertion style', () => {
		const el = textElement([
			seg('◆ ', {
				bulletInfo: { char: '◆' },
				paragraphInsertionStyle: { fontSize: 30, italic: true },
			}),
		]);
		const { patch } = setSelectionBullets(el, 'numbered', null);
		expect(texts(patch)).toStrictEqual(['1.']);
		expect(
			'textSegments' in patch && patch.textSegments?.[0].paragraphInsertionStyle,
		).toStrictEqual({
			fontSize: 30,
			italic: true,
		});
	});

	it('restores custom marker metadata when toggling a selection off and on', () => {
		const el = textElement([
			seg('◆ ', { bulletInfo: { char: '◆', fontFamily: 'Wingdings', sizePercent: 75 } }),
			seg('Body'),
		]);
		const off = toggleSelectionBullets(el, 'bullet', {
			startSegIdx: 1,
			startOffset: 0,
			endSegIdx: 1,
			endOffset: 4,
		});
		const on = toggleSelectionBullets(
			{ ...el, ...off.patch } as PptxElement,
			'bullet',
			off.newSelection,
		);
		expect(texts(on.patch)).toStrictEqual(['◆ ', 'Body']);
		expect('textSegments' in on.patch && on.patch.textSegments?.[0].bulletInfo).toMatchObject({
			char: '◆',
			fontFamily: 'Wingdings',
			sizePercent: 75,
		});
	});

	it('applies the kind to a mixed selection rather than turning it off', () => {
		const el = textElement([...loadedBullet('A'), brk(), seg('B')]);
		const { patch } = toggleSelectionBullets(el, 'bullet', {
			startSegIdx: 1,
			startOffset: 0,
			endSegIdx: 3,
			endOffset: 1,
		});
		const segments = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(segments.filter((s) => isBulletMarkerSegment(s))).toHaveLength(2);
	});

	it('is the whole-element toggle when there is no selection', () => {
		const el = textElement([seg('A'), brk(), seg('B')]);
		const { patch, newSelection } = toggleSelectionBullets(el, 'bullet', null);
		expect(patch).toStrictEqual(toggleElementBullets(el, 'bullet'));
		expect(newSelection).toBeNull();
	});

	it('acts on a live segment override and clears the inert listType', () => {
		const el = {
			...textElement([seg('stale')]),
			textStyle: { fontSize: 18, listType: 'none' },
		} as unknown as PptxElement;
		const { patch } = setSelectionBullets(el, 'bullet', null, [seg('live')]);
		expect(texts(patch)).toStrictEqual(['• ', 'live']);
		expect((patch as { textStyle: { listType?: string } }).textStyle.listType).toBeUndefined();
	});
});

describe("listConversionScope: 'list' switches the whole list around a caret (Word / Google Docs)", () => {
	const LIST = { listConversionScope: 'list' } as const;
	const bullet = (text: string, level = 0): TextSegment[] => [
		seg('•', { bulletInfo: { char: '•' }, paragraphLevel: level }),
		seg(text, { paragraphLevel: level }),
	];
	const numbered = (text: string, ordinal: number, level = 0): TextSegment[] => [
		seg(`${ordinal + 1}.`, {
			bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: ordinal },
			paragraphLevel: level,
		}),
		seg(text, { paragraphLevel: level }),
	];
	const kinds = (patch: Partial<PptxElement>) =>
		splitBulletParagraphs((patch as { textSegments: TextSegment[] }).textSegments).map(
			(paragraph) => paragraphBulletKind(paragraph.segments),
		);
	/** A collapsed caret inside the text run at `index`. */
	const caret = (index: number, offset = 1) => ({
		startSegIdx: index,
		startOffset: offset,
		endSegIdx: index,
		endOffset: offset,
	});

	it('converts every item of a bulleted list to numbered, from any item', () => {
		const el = textElement([...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')]);
		for (const index of [1, 4, 7]) {
			const { patch } = toggleSelectionBullets(el, 'numbered', caret(index), undefined, LIST);
			expect(kinds(patch)).toStrictEqual(['numbered', 'numbered', 'numbered']);
		}
	});

	it('numbers the converted list consecutively', () => {
		const el = textElement([...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')]);
		const { patch } = toggleSelectionBullets(el, 'numbered', caret(4), undefined, LIST);
		const segments = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(
			splitBulletParagraphs(segments).map((p) => p.segments[0]?.bulletInfo?.paragraphIndex),
		).toStrictEqual([0, 1, 2]);
	});

	it('converts a numbered list back to bullets', () => {
		const el = textElement([...numbered('A', 0), brk(), ...numbered('B', 1)]);
		const { patch } = toggleSelectionBullets(el, 'bullet', caret(1), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['bullet', 'bullet']);
	});

	it('keeps the caret where it was', () => {
		const el = textElement([...bullet('A'), brk(), ...bullet('Bee'), brk(), ...bullet('C')]);
		const { patch, newSelection } = toggleSelectionBullets(
			el,
			'numbered',
			caret(4, 2),
			undefined,
			LIST,
		);
		const segments = (patch as { textSegments: TextSegment[] }).textSegments;
		expect(newSelection?.startSegIdx).toBe(newSelection?.endSegIdx);
		expect(newSelection?.startOffset).toBe(newSelection?.endOffset);
		expect(segments[newSelection!.startSegIdx].text).toBe('Bee');
		expect(newSelection?.startOffset).toBe(2);
	});

	it('stops at a paragraph with no list on either side', () => {
		const el = textElement([
			seg('Intro'),
			brk(),
			...bullet('A'),
			brk(),
			...bullet('B'),
			brk(),
			seg('Outro'),
			brk(),
			...bullet('Z'),
		]);
		const { patch } = toggleSelectionBullets(el, 'numbered', caret(3), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['none', 'numbered', 'numbered', 'none', 'bullet']);
	});

	it('stops at a same-level item of the other kind', () => {
		const el = textElement([...numbered('A', 0), brk(), ...bullet('B'), brk(), ...bullet('C')]);
		const { patch } = toggleSelectionBullets(el, 'numbered', caret(4), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['numbered', 'numbered', 'numbered']);
		const back = toggleSelectionBullets(el, 'bullet', caret(1), undefined, LIST);
		expect(kinds(back.patch)).toStrictEqual(['bullet', 'bullet', 'bullet']);
	});

	it('carries nested deeper items with the list, and stops at a shallower level', () => {
		const el = textElement([
			...bullet('Top'),
			brk(),
			...bullet('A', 1),
			brk(),
			...bullet('a1', 2),
			brk(),
			...bullet('B', 1),
			brk(),
			...bullet('Next top'),
		]);
		const { patch } = toggleSelectionBullets(el, 'numbered', caret(4), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['bullet', 'numbered', 'numbered', 'numbered', 'bullet']);
	});

	it('turns a list off for the caret paragraph only', () => {
		const el = textElement([...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')]);
		const { patch } = toggleSelectionBullets(el, 'bullet', caret(4), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['bullet', 'none', 'bullet']);
	});

	it('lists only the caret paragraph when it had no list', () => {
		const el = textElement([seg('A'), brk(), seg('B'), brk(), seg('C')]);
		const { patch } = toggleSelectionBullets(el, 'numbered', caret(2), undefined, LIST);
		expect(kinds(patch)).toStrictEqual(['none', 'numbered', 'none']);
	});

	it('keeps a range selection scoped to the paragraphs it touches', () => {
		const el = textElement([...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')]);
		const { patch } = toggleSelectionBullets(
			el,
			'numbered',
			{ startSegIdx: 4, startOffset: 0, endSegIdx: 4, endOffset: 1 },
			undefined,
			LIST,
		);
		expect(kinds(patch)).toStrictEqual(['bullet', 'numbered', 'bullet']);
	});
});

describe('the default list conversion scope matches PowerPoint', () => {
	const bullet = (text: string): TextSegment[] => [
		seg('•', { bulletInfo: { char: '•' } }),
		seg(text),
	];
	const kinds = (patch: Partial<PptxElement>) =>
		splitBulletParagraphs((patch as { textSegments: TextSegment[] }).textSegments).map(
			(paragraph) => paragraphBulletKind(paragraph.segments),
		);
	const caret = { startSegIdx: 4, startOffset: 1, endSegIdx: 4, endOffset: 1 };
	const el = textElement([...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')]);

	it.each([
		['unset', undefined],
		['paragraph', { listConversionScope: 'paragraph' } as const],
	])('a caret converts only its own paragraph (%s)', (_name, options) => {
		const { patch } = toggleSelectionBullets(el, 'numbered', caret, undefined, options);
		expect(kinds(patch)).toStrictEqual(['bullet', 'numbered', 'bullet']);
	});

	it('the format-panel path agrees', () => {
		const { patch } = applyListStyleUpdate(el, { listType: 'numbered' }, caret);
		expect(kinds(patch)).toStrictEqual(['bullet', 'numbered', 'bullet']);
		const wide = applyListStyleUpdate(el, { listType: 'numbered' }, caret, {
			listConversionScope: 'list',
		});
		expect(kinds(wide.patch)).toStrictEqual(['numbered', 'numbered', 'numbered']);
	});
});
