import { expect, it } from 'vitest';
import { completedNoteStoryMinimum } from './table-note-completion';
import type { RowLayout } from './table-layout';
import type { LayoutParagraphBox } from './result';
const box = (id: string, lines = 1): LayoutParagraphBox => ({
	kind: 'paragraph',
	blockId: id,
	yPx: 5,
	heightPx: 12 * lines,
	tableFrameInsets: { top: 2, bottom: 0.5 },
	lines: Array.from({ length: lines }, (_, i) => ({
		yPx: i * 12,
		heightPx: 12,
		widthPx: 1,
		sourceStart: i,
		sourceEnd: i + 1,
		fragments: [],
	})),
});
const row = (boxes: LayoutParagraphBox[]): RowLayout => ({
	heightPx: 40,
	cells: [boxes],
	geometry: [],
	cellPadding: [{ top: 3, bottom: 3 }],
});
it('reserves shared frame and cell bottom insets for an unreferenced final leaf line', () => {
	expect(
		completedNoteStoryMinimum(
			row([box('a'), box('tail')]),
			new Map([['a', 'inner']]),
			new Set(['outside']),
			100,
		),
	).toBe(20.5);
});
it.each(['multiple', 'pending', 'oversized', 'sequential'] as const)(
	'does not widen the reference band for %s tails',
	(kind) => {
		const boxes =
			kind === 'multiple'
				? [box('a', 2)]
				: kind === 'sequential'
					? [box('a'), box('b')]
					: [box('a')];
		const stories = new Map([
			['a', 'inner'],
			['b', 'inner'],
		]);
		expect(
			completedNoteStoryMinimum(
				row(boxes),
				stories,
				new Set(kind === 'pending' ? ['a'] : []),
				kind === 'oversized' ? 20 : 100,
			),
		).toBe(0);
	},
);
