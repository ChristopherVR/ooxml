import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { layoutParagraph } from './paragraph-layout';
import type { TextMeasurer } from './measure';

// Half an em per character, so every width is exact and no font file is needed.
const measurer: TextMeasurer = {
	widthOf: (text, font) => text.length * font.sizePx * 0.5,
	lineHeightOf: (font) => font.sizePx * 1.15,
};

async function lines(name: string, textMeasurer: TextMeasurer = measurer) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, textMeasurer);
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) => (block.kind === 'paragraph' ? block.lines : [])),
		),
	);
}
type Line = Awaited<ReturnType<typeof lines>>[number];
const words = (line: Line) =>
	line.fragments
		.filter((f) => f.text.trim())
		.map((f) => ({ text: f.text, x: f.xPx, w: f.widthPx }));
/** Where the last fragment of the line ends. */
const end = (line: Line) => {
	const last = line.fragments.at(-1)!;
	return last.xPx + last.widthPx;
};

describe('tab stops inherited from paragraph styles (#414)', () => {
	it('places the tabs at the stops the style defines, as with direct stops', async () => {
		const [style, derived] = await lines('issue-414-style-tab-stops');
		const [direct] = await lines('issue-414-direct-tab-stops');
		expect(words(style!)).toEqual(words(direct!));
		const [, value] = words(style!);
		// 4536 twips = 302.4 px; the right stop at 9070 twips = 604.67 px.
		expect(value!.x).toBeCloseTo(302.4, 1);
		expect(end(style!)).toBeCloseTo(604.67, 1);
		// A derived style clears 4536, adds 2000 and keeps 9070.
		expect(words(derived!)[1]!.x).toBeCloseTo(133.33, 1);
		expect(end(derived!)).toBeCloseTo(604.67, 1);
	});
});

const texts = (lineList: Line[]) =>
	lineList.map((line) => line.fragments.map((f) => f.text).join(''));

describe('a right tab stop past the right margin (#385)', () => {
	it('ends the text at the margin on one line, as with a stop at the margin', async () => {
		const bug = await lines('issue-385-right-tab-past-margin');
		const control = await lines('issue-385-right-tab-at-margin');
		expect(control).toHaveLength(3);
		expect(bug).toHaveLength(3);
		// 9070 twips of text area = 604.67 px.
		for (const line of [...bug, ...control]) expect(end(line)).toBeCloseTo(604.67, 1);
		expect(texts(bug)).toEqual(texts(control));
	});
});

describe('text after a right tab that ends at the right margin (#404)', () => {
	it('keeps every tabbed and positional-tab paragraph on one line', async () => {
		// The widths are exact here: the stop and the summed words differ only in the last bits.
		expect(await lines('issue-404-tab-stop-one-twip-inside')).toHaveLength(5);
		// Five paragraphs with a tab stop at 9072 twips and five with a right `w:ptab`.
		expect(await lines('issue-404-tab-stop-at-margin')).toHaveLength(10);
	});

	// 10 px per character, 'W' 1 pt (1.333 px) wider.
	const wide: TextMeasurer = {
		widthOf: (text) => [...text].reduce((sum, c) => sum + (c === 'W' ? 11.333333 : 10), 0),
		lineHeightOf: () => 20,
	};
	const tabbed = (value: string, align: 'right' | 'center' = 'right') =>
		layoutParagraph(
			{
				kind: 'paragraph',
				id: 'p',
				runs: [{ text: `a	${value}` }],
				tabStops: [{ posPx: 200, align }],
			},
			200,
			wide,
			() => {},
		).lines.map((line) => line.fragments.map((f) => f.text).join(''));

	it('still wraps text that overflows the margin by 1 pt', () => {
		expect(tabbed('b'.repeat(18) + 'W')).toHaveLength(2);
		expect(tabbed('b'.repeat(19))).toEqual(['a' + 'b'.repeat(19)]);
	});
});

describe('a left tab whose following text does not fit (#420)', () => {
	// 1134 and 3000 twips from the margin, 75.6 and 200 px.
	it('moves the tab to the next line, to the first stop', async () => {
		const bug = await lines('issue-420-left-tab-text-does-not-fit');
		expect(texts(bug)).toEqual([
			'Stops at 1134 and 8000 twips',
			'Extraordinarily long words follow',
			'Stops at 3000 and 8000 twips',
			'Extraordinarily long words follow',
		]);
		const second = (line: Line) => words(line).find((word) => word.text === 'Extraordinarily')!;
		expect(second(bug[1]!).x).toBeCloseTo(75.6, 1);
		expect(second(bug[3]!).x).toBeCloseTo(200, 1);
	});

	it('keeps the text on the line when it fits after the last stop', async () => {
		const control = await lines('issue-420-left-tab-text-fits-control');
		expect(control).toHaveLength(2);
		expect(words(control[0]!).at(-1)!.x).toBeCloseTo(533.33, 1);
	});
});

describe('the tab after a list number stops at the hanging indent (#422)', () => {
	// The bullet ends before 384 twips (25.6 px); the custom stops are 200, 340, 500 and 1235 twips.
	it('uses a custom stop only when it comes before the hanging indent', async () => {
		// A bullet of 80 twips, so that it ends at 322 twips as in Word's measurement.
		const bullet: TextMeasurer = {
			widthOf: (text, font) => text.length * font.sizePx * (text === '-' ? 0.3333 : 0.5),
			lineHeightOf: measurer.lineHeightOf,
		};
		const bug = await lines('issue-422-list-tab-past-indent', bullet);
		const control = await lines('issue-422-list-no-custom-tab-control', bullet);
		const start = (line: Line) => words(line).find((word) => word.text.startsWith('Text'))!.x;
		expect(bug).toHaveLength(4);
		const indent = start(control[0]!);
		expect(indent).toBeCloseTo(25.6, 1);
		expect(bug.map(start)).toEqual([
			expect.closeTo(indent, 1),
			expect.closeTo(22.67, 1),
			expect.closeTo(indent, 1),
			expect.closeTo(indent, 1),
		]);
	});
});
