import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import {
	anchor,
	CHART_PARTS,
	cellFloatCases,
	cellFloatPackage,
} from '../__fixtures__/cell-floats/cases';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';
import type { LayoutParagraphBox, LayoutTableBox } from './result';

interface Recorded {
	pages: number;
	objects: { x: number; y: number; width: number; height: number }[];
	rules: number[];
	/** [text, x, y, end x] */
	words: [string, number, number, number][];
}
const fixtures = new URL('../__fixtures__/cell-floats/', import.meta.url);
const recorded = JSON.parse(readFileSync(new URL('word-positions.json', fixtures), 'utf8')) as {
	cases: Record<string, Recorded>;
};
const cases = cellFloatCases();
const chartPackage = unzipSync(
	readFileSync(new URL('../__fixtures__/chart.docx', import.meta.url)),
);
const chartParts = Object.fromEntries(
	CHART_PARTS.map((part) => [part, strFromU8(chartPackage[part]!)]),
);

/** Courier New, the fixtures' only font: every character advances 1229/2048 em. */
const courier: TextMeasurer = {
	widthOf: (text, font) =>
		[...text].filter((ch) => ch !== '\t' && ch !== '\n').length * font.sizePx * (1229 / 2048),
	// Word lays Courier New 11 pt lines 12.5 pt apart.
	lineHeightOf: (font) => font.sizePx * (12.5 / 11),
};
const pt = (px: number) => Math.round(((px * 72) / 96) * 100) / 100;

async function layOut(name: string) {
	const entry = cases[name]!;
	const parts = cellFloatPackage(entry, chartParts);
	const bytes = zipSync(
		Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, strToU8(xml)])),
	);
	const { model } = await loadDocx(bytes);
	const result = layoutDocumentModel(model, courier);
	const page = result.pages[0]!;
	const column = page.columns[0]!;
	const table = column.blocks.find((block): block is LayoutTableBox => block.kind === 'table');
	const float = page.floats?.[0];
	return {
		result,
		page,
		float: float && { ...float, x: pt(float.xPx), y: pt(float.yPx) },
		/** The first row's bottom edge (pt from the page top). */
		rowBottom: table && pt(page.marginTopPx + table.yPx + table.rows[0]!.heightPx),
		/** The page x of the anchor cell's text area. */
		cellTextLeft:
			table &&
			pt(
				page.marginLeftPx +
					column.xPx +
					(table.xPx ?? 0) +
					table.rows[0]!.geometry![1]!.xPx +
					table.rows[0]!.geometry![1]!.paddingLeftPx,
			),
		/** Lines of the anchor cell's last paragraph as [x, y] in points. */
		cellLines: () => {
			const row = table!.rows[0]!;
			const geometry = row.geometry![1]!;
			const box = row.cells[1]!.at(-1) as LayoutParagraphBox;
			const left = page.marginLeftPx + column.xPx + (table!.xPx ?? 0) + geometry.xPx;
			return box.lines.map((line) => [
				pt(left + geometry.paddingLeftPx + (line.fragments[0]?.xPx ?? 0)),
				pt(page.marginTopPx + table!.yPx + row.yPx + box.yPx + line.yPx),
			]);
		},
	};
}

/** Word's lines (by y) of the words right of `minX`, as [x, y] of each line's first word. */
function wordLines(word: Recorded, minX: number, minY: number): number[][] {
	const lines = new Map<number, number>();
	for (const [, x, y] of word.words)
		if (x >= minX && y >= minY) lines.set(y, Math.min(lines.get(y) ?? Infinity, x));
	return [...lines].sort((a, b) => a[0] - b[0]).map(([y, x]) => [x, y]);
}

const inCellCases = Object.keys(recorded.cases).filter(
	(name) => /^(h|v|grow|wrap)-|^behind$/.test(name) && name !== 'h-character-body',
);

describe('floating objects in table cells against Word 16', () => {
	it('records every case it builds', () => {
		expect(Object.keys(recorded.cases).sort()).toEqual(Object.keys(cases).sort());
	});

	it.each(inCellCases)('places %s where Word draws it', async (name) => {
		const word = recorded.cases[name]!;
		const { float } = await layOut(name);
		const expected = word.objects[0]!;
		expect(float, name).toBeDefined();
		expect(Math.abs(float!.x - expected.x), `${name} x ${float!.x}`).toBeLessThanOrEqual(1);
		expect(Math.abs(float!.y - expected.y), `${name} y ${float!.y}`).toBeLessThanOrEqual(1);
		expect(float!.inCell).toBe(true);
	});

	it.each(inCellCases)('grows the row of %s as Word does', async (name) => {
		const word = recorded.cases[name]!;
		const { rowBottom } = await layOut(name);
		expect(Math.abs(rowBottom! - word.rules[1]!), `${name} row ${rowBottom}`).toBeLessThanOrEqual(
			1,
		);
	});

	it.each(['square', 'largest', 'tight', 'topAndBottom', 'none'])(
		'wraps the cell text around a %s object as Word does',
		async (wrap) => {
			for (const name of [`wrap-${wrap}`, `wrap-right-${wrap}`]) {
				const word = wordLines(recorded.cases[name]!, 320, 120).slice(0, 6);
				const lines = (await layOut(name)).cellLines();
				for (const [index, [x, y]] of word.entries()) {
					const [lineX, lineY] = lines[index]!;
					expect(Math.abs(lineX! - x!), `${name} line ${index} x ${lineX}`).toBeLessThanOrEqual(1);
					expect(Math.abs(lineY! - y!), `${name} line ${index} y ${lineY}`).toBeLessThanOrEqual(
						1.5,
					);
				}
			}
		},
	);

	it('ignores layoutInCell="0" from compatibility mode 15, as Word 2013 and later do', async () => {
		for (const name of ['h-page-lic0-square', 'v-margin-lic0-none']) {
			const { float, result } = await layOut(name);
			expect(float?.inCell, name).toBe(true);
			expect(result.pages[0]!.floats).toHaveLength(1);
		}
	});

	it.each(Object.keys(recorded.cases).filter((name) => /^m14-.*-lic0-/.test(name)))(
		'positions %s on the page, as Word 2010 does with layoutInCell="0"',
		async (name) => {
			const expected = recorded.cases[name]!.objects[0]!;
			const { float } = await layOut(name);
			expect(float?.inCell, name).toBeUndefined();
			expect(Math.abs(float!.x - expected.x), `${name} x ${float!.x}`).toBeLessThanOrEqual(1);
			expect(Math.abs(float!.y - expected.y), `${name} y ${float!.y}`).toBeLessThanOrEqual(1);
		},
	);

	it.each(Object.keys(recorded.cases).filter((name) => /^m14-.*-lic1-/.test(name)))(
		'lays %s out in its cell, as Word 2010 does with layoutInCell="1"',
		async (name) => {
			const word = recorded.cases[name]!;
			const expected = word.objects[0]!;
			// Word 2010 outdents the table by its cell margin, which the layout does not model: compare
			// the object's distance from the anchor cell's text.
			const unwrapped = recorded.cases[name.replace(/-square$/, '-none')]!;
			const anchorX = unwrapped.words.find(([text]) => text === 'Anchor')![1];
			const { float, cellTextLeft } = await layOut(name);
			expect(float?.inCell, name).toBe(true);
			expect(Math.abs(float!.x - cellTextLeft! - (expected.x - anchorX))).toBeLessThanOrEqual(1);
		},
	);

	it('measures a body paragraph character and line frame from the character before the anchor', async () => {
		const expected = recorded.cases['h-character-body']!.objects[0]!;
		const { float } = await layOut('h-character-body');
		expect(float?.inCell).toBeUndefined();
		expect(Math.abs(float!.x - expected.x)).toBeLessThanOrEqual(1);
		expect(Math.abs(float!.y - expected.y)).toBeLessThanOrEqual(1);
	});
});

describe('floating charts against Word 16', () => {
	it.each(['chart-body', 'chart-body-topAndBottom', 'chart-cell'])(
		'places %s where Word draws it, painted like an inline chart',
		async (name) => {
			const expected = recorded.cases[name]!.objects[0]!;
			const { float } = await layOut(name);
			expect(float?.chart?.svg, name).toContain('<svg');
			expect(float!.chart!.shape).toBeUndefined();
			expect(Math.abs(float!.x - expected.x), `${name} x ${float!.x}`).toBeLessThanOrEqual(1);
			expect(Math.abs(float!.y - expected.y), `${name} y ${float!.y}`).toBeLessThanOrEqual(1);
			expect(pt(float!.widthPx)).toBeCloseTo(expected.width, 0);
			expect(pt(float!.heightPx)).toBeCloseTo(expected.height, 0);
		},
	);

	it('wraps the body text beside a square-wrapped chart', async () => {
		const { page } = await layOut('chart-body');
		const word = recorded.cases['chart-body']!;
		// Word starts the lines beside the chart 9 pt right of it (the chart spans 144 to 324 pt).
		const beside = word.words.filter(([, x, y]) => x > 330 && y > 115 && y < 235);
		expect(beside.length).toBeGreaterThan(0);
		const resumeX = Math.min(...beside.map(([, x]) => x));
		const paragraph = page.columns[0]!.blocks.find(
			(block): block is LayoutParagraphBox => block.kind === 'paragraph' && block.lines.length > 3,
		)!;
		const starts = paragraph.lines
			.filter((line) => pt(page.marginTopPx + paragraph.yPx + line.yPx) < 230)
			.flatMap((line) => line.fragments.map((fragment) => pt(page.marginLeftPx + fragment.xPx)))
			.filter((x) => x > 200);
		expect(starts.length).toBeGreaterThan(0);
		expect(Math.abs(Math.min(...starts) - resumeX)).toBeLessThanOrEqual(1);
	});

	it('moves the body text below a top-and-bottom chart', async () => {
		const { page, float } = await layOut('chart-body-topAndBottom');
		const word = recorded.cases['chart-body-topAndBottom']!;
		const bottom = float!.y + pt(float!.heightPx);
		// Word's first line below the chart.
		const below = Math.min(...word.words.filter(([, , y]) => y >= bottom - 1).map(([, , y]) => y));
		const lineTops = page.columns[0]!.blocks.flatMap((block) =>
			block.kind === 'paragraph'
				? block.lines.map((line) => pt(page.marginTopPx + block.yPx + line.yPx))
				: [],
		);
		expect(lineTops.some((y) => y > float!.y - 12 && y < bottom - 1)).toBe(false);
		expect(Math.min(...lineTops.filter((y) => y >= bottom - 1)) - below).toBeLessThanOrEqual(1.5);
	});
});

describe('text boxes in table cells against Word 16', () => {
	it('lays a text box in a cell out in the cell, its lines broken by the page engine as Word breaks them', async () => {
		const word = recorded.cases['textbox-cell']!;
		const expected = word.objects[0]!;
		const { float, rowBottom } = await layOut('textbox-cell');
		expect(float?.inCell).toBe(true);
		expect(Math.abs(float!.x - expected.x)).toBeLessThanOrEqual(1);
		expect(Math.abs(float!.y - expected.y)).toBeLessThanOrEqual(1);
		expect(Math.abs(rowBottom! - word.rules[1]!)).toBeLessThanOrEqual(1);
		// Word's lines inside the box, by their tops; the line cut by the box's bottom is clipped.
		const bottom = expected.y + expected.height;
		const inside = new Map<number, string[]>();
		for (const [text, x, y] of word.words)
			if (x >= expected.x && x < expected.x + expected.width && y >= expected.y && y < bottom)
				inside.set(y, [...(inside.get(y) ?? []), text]);
		const wordLines = [...inside]
			.sort((a, b) => a[0] - b[0])
			.map(([, texts]) => texts.join(' '))
			.slice(0, 4);
		const layout = float!.textBoxLayout!;
		const engineLines = layout.blocks.flatMap((block) =>
			(block as LayoutParagraphBox).lines.map((line) =>
				line.fragments
					.map((fragment) => fragment.text)
					.join('')
					.trim(),
			),
		);
		expect(engineLines.slice(0, 4)).toEqual(wordLines);
		// More lines than the box shows: the rest is clipped at the text rectangle.
		const lastBottom = Math.max(
			...layout.blocks.flatMap((block) =>
				(block as LayoutParagraphBox).lines.map((line) => block.yPx + line.yPx + line.heightPx),
			),
		);
		expect(lastBottom).toBeGreaterThan(layout.frame.heightPx);
	});
});

describe('character frames under each review display mode', () => {
	// "Keep " + a deletion "GONE " + an insertion "NEW " + "here", then a character-anchored shape:
	// the frame starts at the last character shown before the anchor, whichever runs a mode hides.
	const body =
		'<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r>' +
		'<w:del w:id="1" w:author="A"><w:r><w:delText xml:space="preserve">GONE </w:delText></w:r></w:del>' +
		'<w:ins w:id="2" w:author="A"><w:r><w:t xml:space="preserve">NEW </w:t></w:r></w:ins>' +
		`<w:r><w:t>here</w:t></w:r>${anchor(1, 'character:18', 'line:0', 'none', 1)}</w:p>`;
	const shown = {
		all: 'Keep GONE NEW here',
		simple: 'Keep NEW here',
		final: 'Keep NEW here',
		original: 'Keep GONE here',
	};

	it.each(Object.entries(shown))('measures from the shown text in %s', async (mode, text) => {
		const parts = cellFloatPackage({ body, mode: 15 });
		const bytes = zipSync(
			Object.fromEntries(Object.entries(parts).map(([path, xml]) => [path, strToU8(xml)])),
		);
		const { model } = await loadDocx(bytes);
		const result = layoutDocumentModel(model, courier, {
			reviewDisplayMode: mode as 'all' | 'simple' | 'final' | 'original',
		});
		const page = result.pages[0]!;
		const float = page.floats![0]!;
		const advance = 11 * (96 / 72) * (1229 / 2048);
		const expected = page.marginLeftPx + (text.length - 1) * advance + 24;
		expect(Math.abs(float.xPx - expected), `${mode} x ${float.xPx}`).toBeLessThanOrEqual(0.5);
	});
});
