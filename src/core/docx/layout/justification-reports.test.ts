import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { TextMeasurer } from './measure';

// Relative Century Gothic advances measured with Pillow at 1000 px on Windows.
// Store only the advances used by the fixtures so the test needs no installed font.
const advance: Record<string, number> = {
	' ': 0.277,
	'.': 0.277,
	A: 0.74,
	C: 0.813,
	D: 0.744,
	E: 0.536,
	I: 0.226,
	L: 0.462,
	M: 0.919,
	O: 0.869,
	P: 0.592,
	R: 0.607,
	T: 0.426,
	V: 0.702,
	a: 0.683,
	b: 0.682,
	c: 0.647,
	d: 0.685,
	e: 0.65,
	f: 0.314,
	g: 0.673,
	h: 0.61,
	i: 0.2,
	l: 0.2,
	m: 0.938,
	n: 0.61,
	o: 0.655,
	p: 0.682,
	q: 0.682,
	r: 0.301,
	s: 0.388,
	t: 0.339,
	u: 0.608,
	v: 0.554,
	x: 0.48,
};
const measurer: TextMeasurer = {
	widthOf(text, font) {
		return (
			[...text].reduce((sum, char) => {
				const width = advance[char];
				if (width === undefined) throw new Error(`Missing fixture advance for ${char}`);
				return sum + width;
			}, 0) * font.sizePx
		);
	},
	lineHeightOf: (font) => font.sizePx * 1.227,
};

it.each([
	{ mode: 15, name: 'justify-compat15', pageCount: 2, endings: ['aliqua', 'aute', 'velit'] },
	{ mode: 14, name: 'justify-compat14-control', pageCount: 3, endings: ['', '', ''] },
])(
	'wraps the reported lines in compatibility mode $mode (#346)',
	async ({ mode, name, pageCount, endings }) => {
		const { model } = await loadDocx(
			await readFile(
				new URL(`./fixtures/word-layout-reports/issue-346-${name}.docx`, import.meta.url),
			),
		);
		expect(model.compatibilityMode).toBe(mode);
		const { pages } = layoutDocumentModel(model, measurer);
		const lines = pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph'
						? block.lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
						: [],
				),
			),
		);
		const prefixes = [
			'proident enim do aliqua exercitation',
			'pariatur irure adipiscing dolor in labore',
			'magna fugiat enim culpa eiusmod sunt',
		];
		for (const [index, prefix] of prefixes.entries())
			expect(lines.find((line) => line.startsWith(prefix))).toBe(
				`${prefix} ${endings[index] ?? ''}`.trimEnd(),
			);
		expect(pages).toHaveLength(pageCount);
	},
);
