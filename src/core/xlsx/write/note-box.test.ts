import { describe, expect, it } from 'vitest';
import { createGridMetrics } from '../layout/metrics';
import { createWorkbook } from '../workbook';
import type { Comment } from '../model';
import { commentParts, PersonRegistry } from './comments';
import { NOTE_MIN_HEIGHT_PX, NOTE_WIDTH_PX, noteBox, noteLines } from './note-box';

const note = (row: number, col: number, text: string): Comment =>
	({ address: { row, col }, author: 'Ana', text }) as Comment;

describe('legacy note boxes', () => {
	const sheet = createWorkbook().sheets[0]!;
	const grid = createGridMetrics(sheet, { zoom: 100 });

	it('keeps the default box for a short note and grows down for a long one', () => {
		const short = noteBox(grid, 3, 2, 'Check');
		expect(short.style).toBe(
			`margin-left:${(grid.colLeft(3) + 15) * 0.75}pt;margin-top:${(grid.rowTop(2) + 10) * 0.75}pt;width:108pt;height:59.25pt`,
		);
		const text = 'A longer note that runs over several lines of the box. '.repeat(4);
		const long = noteBox(grid, 3, 2, text);
		const height = Number(/height:([\d.]+)pt/.exec(long.style)?.[1]) / 0.75;
		expect(noteLines(text)).toBeGreaterThan(5);
		expect(height).toBeGreaterThan(NOTE_MIN_HEIGHT_PX);
		const [left, , top, , right, , bottom] = long.anchor.split(', ').map(Number);
		expect([left, top]).toEqual([3, 2]);
		expect(grid.colLeft(right!) - grid.colLeft(3)).toBeLessThanOrEqual(NOTE_WIDTH_PX + 15);
		expect(grid.rowTop(bottom!) + 1).toBeGreaterThan(grid.rowTop(2) + height - 20);
		expect(bottom).toBeGreaterThan(Number(short.anchor.split(', ')[6]));
	});

	it('sizes each legacy note in the VML drawing, not a threaded comment', () => {
		const long = 'Line one\nLine two\nLine three\nLine four\nLine five\nLine six';
		const { vml } = commentParts(
			[note(0, 0, long), { ...note(5, 0, 'Thread'), replies: [] }],
			1,
			new PersonRegistry(),
			'Sheet1',
			undefined,
			undefined,
			grid,
		);
		const heights = [...vml.matchAll(/height:([\d.]+)pt/g)].map((match) => Number(match[1]));
		expect(heights[0]).toBeGreaterThan(59.25);
		expect(heights[1]).toBe(59.25);
		expect(vml).toContain('<x:Anchor>1, 15, 0, 10, ');
	});
});
