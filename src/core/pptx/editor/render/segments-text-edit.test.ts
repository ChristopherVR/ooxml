import type { TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { editSegmentsText, segmentsPlainText } from './segments-text-edit';
import { cellRunsFromSegments, tableCellSegments, withCellSegments } from './table-cell-rich-text';
import { withCellTextStyle } from './table-cells-text-style';

const runs: TextSegment[] = [
	{ text: 'Revenue ', style: { fontSize: 20 }, paragraphProperties: { align: 'center' } },
	{ text: 'grew', style: { bold: true } },
	{ text: '\n', isParagraphBreak: true, style: {} },
	{ text: 'Second', style: { italic: true } },
];

const shape = (segments: TextSegment[]) =>
	segments.map((s) => [s.isParagraphBreak ? '<P>' : s.text, Boolean(s.style.bold)]);

describe('editSegmentsText', () => {
	it('replaces one word without moving the run boundaries around it', () => {
		const next = editSegmentsText(runs, 'Income grew\nSecond', 6);
		expect(shape(next)).toEqual([
			['Income ', false],
			['grew', true],
			['<P>', false],
			['Second', false],
		]);
		expect(next[0]!.style.fontSize).toBe(20);
		expect(next[0]!.paragraphProperties).toEqual({ align: 'center' });
	});

	it('gives typed text the format of the character before it', () => {
		const next = editSegmentsText(runs, 'Revenue grew!\nSecond', 12);
		expect(next[1]).toMatchObject({ text: 'grew!', style: { bold: true } });
	});

	it('splits a paragraph on an inserted break, keeping its layout on both halves', () => {
		const next = editSegmentsText(runs, 'Revenue \ngrew\nSecond', 9);
		expect(segmentsPlainText(next)).toBe('Revenue \ngrew\nSecond');
		const breaks = next.filter((s) => s.isParagraphBreak);
		expect(breaks).toHaveLength(2);
		expect(next[2]).toMatchObject({ text: 'grew', paragraphProperties: { align: 'center' } });
	});

	it('merges paragraphs when a break is deleted', () => {
		const next = editSegmentsText(runs, 'Revenue grewSecond', 12);
		expect(next.some((s) => s.isParagraphBreak)).toBe(false);
		expect(next.at(-1)).toMatchObject({ text: 'Second', style: { italic: true } });
	});

	it('leaves the segments as they are when the text did not change', () => {
		expect(editSegmentsText(runs, segmentsPlainText(runs))).toEqual(runs);
	});
});

describe('table cell rich text', () => {
	it('reads loaded runs as segments with sizes in px and paragraph layout', () => {
		const segments = tableCellSegments({
			text: 'A B\nC',
			textRuns: [
				{ text: 'A ', fontSize: 12 },
				{ text: 'B', bold: true },
				{ text: '', isParagraphBreak: true },
				{ text: 'C', color: '#FF0000' },
			],
			paragraphs: [{ align: 'right' }, {}],
		});
		expect(segments[0]).toMatchObject({
			text: 'A ',
			style: { fontSize: 16 },
			paragraphProperties: { align: 'right' },
		});
		expect(segments[1]).toMatchObject({ text: 'B', style: { bold: true } });
		expect(segments[2]!.isParagraphBreak).toBe(true);
		expect(segments[3]).toMatchObject({ text: 'C', style: { color: '#FF0000' } });
	});

	it('reads a plain cell in its cell style', () => {
		const segments = tableCellSegments({ text: 'x\n', style: { bold: true, fontSize: 9 } });
		expect(segments.map((s) => s.text)).toEqual(['x', '\n', '']);
		expect(segments[0]!.style).toEqual({ bold: true, fontSize: 12 });
	});

	it('keeps text, segments, runs and paragraphs in step', () => {
		const cell = withCellSegments({ text: 'old', paragraphs: [{ align: 'left' }] }, runs);
		expect(cell.text).toBe('Revenue grew\nSecond');
		expect(cell.textSegments).toBe(runs);
		expect(cell.textRuns).toEqual(cellRunsFromSegments(runs));
		expect(cell.textRuns![0]!.fontSize).toBe(15);
		expect(cell.paragraphs).toEqual([{ align: 'center' }, {}]);
	});

	it('keeps a format turned off in an edit, so the cell style cannot paint it back', () => {
		const cell = withCellSegments({ text: 'x', style: { bold: true } }, [
			{ text: 'x', style: { bold: false } },
		]);
		expect(cell.textRuns).toEqual([{ text: 'x', bold: false }]);
	});

	it('drops a theme colour reference when a colour is picked for the cells', () => {
		const cell = withCellTextStyle(
			{ text: 'x', style: { color: '#000000', colorRef: { scheme: 'accent1' } as never } },
			{ color: '#FF0000' },
		);
		expect(cell.style).toMatchObject({ color: '#FF0000', colorRef: undefined });
	});
});
