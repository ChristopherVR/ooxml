import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';
import type { LayoutParagraphBox, LayoutTableBox } from './result';
import { FLOATING_TABLE_NOTE } from './flow-floating-table';
import { floatingFrame } from './adapt-table-floating';

const fixtures = new URL('../__fixtures__/floating-table/', import.meta.url);
interface Recorded {
	cellText: { x: number; y: number; page: number };
	lines: { page: number; x: number; y: number }[];
}
const recorded = JSON.parse(readFileSync(new URL('word-positions.json', fixtures), 'utf8')) as {
	cases: Record<string, Recorded>;
};

/** Courier New, the fixtures' only font: every character advances 1229/2048 em. */
const courier: TextMeasurer = {
	widthOf: (text, font) =>
		[...text].filter((ch) => ch !== '\t' && ch !== '\n').length * font.sizePx * (1229 / 2048),
	// Word lays Courier New 11 pt lines 12.5 pt apart.
	lineHeightOf: (font) => font.sizePx * (12.5 / 11),
};
const pt = (px: number) => (px * 72) / 96;
/** The fixtures' cell margin (108 twips) between the table edge and the cell text. */
const CELL_MARGIN_PT = 5.4;

async function layOut(name: string) {
	const { model } = await loadDocx(new Uint8Array(readFileSync(new URL(`${name}.docx`, fixtures))));
	const result = layoutDocumentModel(model, courier);
	const page = result.pages[0]!;
	const column = page.columns[0]!;
	const table = column.blocks.find((block): block is LayoutTableBox => block.kind === 'table')!;
	const paragraphs = column.blocks.filter(
		(block): block is LayoutParagraphBox => block.kind === 'paragraph',
	);
	const last = paragraphs.at(-1)!;
	const lines = last.lines.map((line) => ({
		x: pt(page.marginLeftPx + column.xPx + (line.fragments[0]?.xPx ?? 0)),
		y: pt(page.marginTopPx + last.yPx + line.yPx),
		fragments: line.fragments.map((fragment) => pt(page.marginLeftPx + fragment.xPx)),
	}));
	return {
		result,
		tableLeft: pt(page.marginLeftPx + column.xPx + (table.xPx ?? 0)),
		tableTop: pt(page.marginTopPx + table.yPx),
		lines,
	};
}

describe('floating tables against Word 16', () => {
	it.each(Object.keys(recorded.cases))('places %s at its frame position', async (name) => {
		const word = recorded.cases[name]!;
		const { tableLeft, tableTop, result } = await layOut(name);
		expect(Math.abs(tableLeft - (word.cellText.x - CELL_MARGIN_PT))).toBeLessThanOrEqual(1);
		// The first cell's text sits half a point below the table top (the top border).
		expect(Math.abs(tableTop - (word.cellText.y - 0.5))).toBeLessThanOrEqual(1);
		expect(result.approximations).toContain(FLOATING_TABLE_NOTE);
	});

	it.each(['margin-left', 'margin-middle', 'page-right', 'margin-center', 'defaults', 'distances'])(
		'wraps the following text around %s',
		async (name) => {
			const word = recorded.cases[name]!.lines.filter((line) => line.page === 1).slice(0, 8);
			const { lines } = await layOut(name);
			for (const [index, expected] of word.entries()) {
				const line = lines[index]!;
				expect(Math.abs(line.y - expected.y), `line ${index} y`).toBeLessThanOrEqual(1.5);
				expect(Math.abs(line.x - expected.x), `line ${index} x`).toBeLessThanOrEqual(1);
			}
		},
	);

	it('continues lines on the far side of a table in the middle of the column', async () => {
		const { lines } = await layOut('margin-middle');
		// The table spans 222-372 pt with 9 pt of distance on each side; the line resumes past it.
		const resumed = lines[0]!.fragments.find((x) => x > 300);
		expect(resumed).toBeGreaterThanOrEqual(372 + 9 - 0.5);
		expect(lines[0]!.fragments.every((x) => x < 222 - 9 || x >= 381 - 0.5)).toBe(true);
	});

	it('does not move the text that follows it', async () => {
		const { lines } = await layOut('margin-left');
		// The paragraph after the table starts where the table starts: the table takes no room.
		expect(Math.abs(lines[0]!.y - 84.5)).toBeLessThanOrEqual(1);
	});
});

describe('floatingFrame', () => {
	it('defaults to the top margin vertically and the column horizontally', () => {
		expect(floatingFrame({})).toEqual({
			relativeFromH: 'column',
			offsetXPx: 0,
			relativeFromV: 'margin',
			offsetYPx: 0,
			distance: { topPx: 0, rightPx: 0, bottomPx: 0, leftPx: 0 },
		});
	});

	it('maps text anchors and inline placement to the following paragraph', () => {
		expect(floatingFrame({ verticalAnchor: 'text', horizontalAnchor: 'text' })).toMatchObject({
			relativeFromH: 'column',
			relativeFromV: 'paragraph',
		});
		const inline = floatingFrame({ verticalAnchor: 'page', yAlign: 'inline' });
		expect(inline.relativeFromV).toBe('page');
	});
});
