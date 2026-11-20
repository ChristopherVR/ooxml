import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel, type InlineImage } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';
import type { LayoutFontSpec, TextMeasurer } from './measure';
import type { LayoutFloatBox, LayoutParagraphBox } from './result';
import type { TextBoxLayout } from './text-box-layout';
import { must } from '../test-support/access';
import {
	BOXES,
	textBoxDocumentXml,
	textBoxLayoutFixture,
	words,
	type FixtureBox,
} from '../test-support/text-box-layout-fixture';

// Text boxes laid out by the page engine (`text-box-layout.ts`), against what Word 16 draws for the
// same `wps:bodyPr` settings (its PDF of `text-box-layout-fixture.ts`): lines wrap inside the text
// rectangle, a linked box takes the lines its first box cannot fit, `spAutoFit` and `wrap="none"`
// resize the shape, `vert` and `vert270` turn the text, and `numCol` is drawn in one column.

const measurer = createFakeMeasurer();
const PX_PER_INCH = 96;
/** Insets (0.1 inch left and right, 0.05 inch top and bottom) and half the 0.75 pt (1 px) outline. */
const PAD = { x: 9.6 + 0.5, y: 4.8 + 0.5 };

async function floats(boxes: FixtureBox[] = BOXES): Promise<Map<string, LayoutFloatBox>> {
	const { model } = await loadDocx(await textBoxLayoutFixture(boxes));
	const result = layoutDocumentModel(model, measurer);
	const placed = result.pages.flatMap((page) => page.floats ?? []);
	return new Map(boxes.map((box, index) => [box.name, must(placed[index], box.name)]));
}
const layoutOf = (float: LayoutFloatBox | undefined): TextBoxLayout =>
	must(float?.textBoxLayout, 'text box layout');
const lines = (layout: TextBoxLayout) =>
	layout.blocks.flatMap((block) =>
		(block as LayoutParagraphBox).lines.map((line) => ({
			text: line.fragments.map((fragment) => fragment.text).join(''),
			top: block.yPx + line.yPx,
			bottom: block.yPx + line.yPx + line.heightPx,
			right: Math.max(0, ...line.fragments.map((fragment) => fragment.xPx + fragment.widthPx)),
		})),
	);
const text = (layout: TextBoxLayout) =>
	lines(layout)
		.map((line) => line.text.trim())
		.join(' ');
const box = (name: string) =>
	must(
		BOXES.find((item) => item.name === name),
		name,
	);

describe('text box layout by the page engine', () => {
	it('reads the wps:bodyPr settings and the chain each box belongs to', async () => {
		const { model } = await loadDocx(await textBoxLayoutFixture());
		const images = drawings(model);
		expect(images.map((image) => image.textBoxContent?.vertical)).toContain('vert270');
		expect(images[4]?.textBoxContent).toMatchObject({ autofit: 'shape' });
		expect(images[5]?.textBoxContent).toMatchObject({ wrap: 'none' });
		expect(images[6]?.textBoxContent?.link).toEqual({ id: 1, seq: 0 });
		expect(images[7]?.textBoxContent).toEqual(
			expect.objectContaining({ blocks: [], link: { id: 1, seq: 1 } }),
		);
		expect(images[0]?.textBoxContent).not.toHaveProperty('link');
	});

	it('wraps lines inside the text rectangle and keeps the overflow below it to be clipped', async () => {
		const layout = layoutOf((await floats()).get('Wrapped'));
		const { width, height } = box('Wrapped');
		expect(layout.frame.xPx).toBeCloseTo(PAD.x, 6);
		expect(layout.frame.yPx).toBeCloseTo(PAD.y, 6);
		expect(layout.frame.widthPx).toBeCloseTo(width * PX_PER_INCH - 2 * PAD.x, 6);
		expect(layout.frame.heightPx).toBeCloseTo(height * PX_PER_INCH - 2 * PAD.y, 6);
		expect(layout).toMatchObject({ widthPx: width * PX_PER_INCH, heightPx: height * PX_PER_INCH });
		const laid = lines(layout);
		expect(laid.length).toBeGreaterThan(5);
		for (const line of laid) expect(line.right).toBeLessThanOrEqual(layout.frame.widthPx + 1e-6);
		expect(text(layout)).toBe(words('Wrapped', 40));
		expect(laid.at(-1)!.top).toBeGreaterThan(layout.frame.heightPx);
	});

	it('flows the text a linked box cannot fit into the next box, at that box’s width', async () => {
		const placed = await floats();
		const first = layoutOf(placed.get('Linked first'));
		const second = layoutOf(placed.get('Linked second'));
		const firstLines = lines(first);
		const fits = firstLines.at(-1)!.bottom;
		expect(fits).toBeLessThanOrEqual(first.frame.heightPx);
		// The next line would not have fitted: it went to the second box.
		expect(fits + (firstLines[0]!.bottom - firstLines[0]!.top)).toBeGreaterThan(
			first.frame.heightPx,
		);
		expect(`${text(first)} ${text(second)}`).toBe(words('Linked', 45));
		expect(lines(second)[0]!.top).toBe(0);
		expect(Math.max(...lines(second).map((line) => line.right))).toBeGreaterThan(
			first.frame.widthPx,
		);
		// The box ahead of the chain's last keeps no overflow; the last keeps it, to be clipped.
		expect(second.frame.widthPx).toBeCloseTo(
			box('Linked second').width * PX_PER_INCH - 2 * PAD.x,
			6,
		);
	});

	it('flows a three-box chain in seq order whatever the boxes’ document order', async () => {
		const chain: FixtureBox[] = [
			{ ...box('Linked second'), id: 1, name: 'Third', content: { linked: 4, seq: 2 }, y: 1 },
			{
				...box('Linked first'),
				id: 2,
				name: 'First',
				chain: 4,
				content: box('Linked first').content,
			},
			{
				...box('Linked second'),
				id: 3,
				name: 'Second',
				content: { linked: 4, seq: 1 },
				height: 0.5,
			},
		];
		const placed = await floats(chain);
		const [third, first, second] = ['Third', 'First', 'Second'].map((name) =>
			text(layoutOf(placed.get(name))),
		);
		expect([first, second, third].join(' ')).toBe(words('Linked', 45));
		expect(second).not.toBe('');
		expect(first!.startsWith('Linked word1')).toBe(true);
	});

	it('turns vert and vert270 text a quarter, with lines along the box’s height', async () => {
		const placed = await floats();
		const down = layoutOf(placed.get('Vertical'));
		const up = layoutOf(placed.get('Vertical 270'));
		expect(down.rotation).toBe(90);
		expect(up.rotation).toBe(270);
		const length = Math.round(box('Vertical').height * PX_PER_INCH) - 2 * PAD.y;
		expect(down.frame.heightPx).toBeCloseTo(length, 6);
		for (const line of lines(down)) expect(line.right).toBeLessThanOrEqual(length + 1e-6);
		// Lines stack across the box's width: three or more lines of a 0.9 inch wide box.
		expect(lines(down).length).toBeGreaterThan(1);
		expect(text(down)).toBe(words('Vertical', 12));
	});

	it('grows a spAutoFit box to its text and leaves normAutofit and noAutofit boxes alone', async () => {
		const placed = await floats();
		const fitted = layoutOf(placed.get('Autofit'));
		const used = lines(fitted).at(-1)!.bottom;
		expect(fitted.frame.heightPx).toBeCloseTo(used, 6);
		expect(fitted.heightPx).toBeCloseTo(used + 2 * PAD.y, 6);
		expect(fitted.heightPx).toBeGreaterThan(box('Autofit').height * PX_PER_INCH);
		// The float is placed, and its shape painted, at the fitted size.
		expect(placed.get('Autofit')?.heightPx).toBe(fitted.heightPx);
		expect(placed.get('Autofit')?.chart?.svg).toContain(`height="${+fitted.heightPx.toFixed(2)}"`);
		const shrink = layoutOf(placed.get('Shrink'));
		expect(shrink.heightPx).toBe(Math.round(box('Shrink').height * PX_PER_INCH));
	});

	it('shrinks a spAutoFit box whose text is shorter than its extent, as Word does', async () => {
		const placed = await floats([
			{ ...box('Autofit'), height: 1.5, content: '<w:p><w:r><w:t>Two words</w:t></w:r></w:p>' },
		]);
		const fitted = layoutOf(placed.get('Autofit'));
		expect(lines(fitted)).toHaveLength(1);
		expect(fitted.heightPx).toBeCloseTo(lines(fitted)[0]!.bottom + 2 * PAD.y, 6);
	});

	it('sizes a wrap="none" box to its longest line', async () => {
		const layout = layoutOf((await floats()).get('No wrap'));
		const laid = lines(layout);
		expect(laid).toHaveLength(1);
		expect(layout.frame.widthPx).toBeCloseTo(laid[0]!.right, 6);
		expect(layout.widthPx).toBeCloseTo(laid[0]!.right + 2 * PAD.x, 6);
		expect(layout.heightPx).toBe(Math.round(box('No wrap').height * PX_PER_INCH));
	});

	it('draws numCol in one column across the box, as Word draws a text box', async () => {
		const layout = layoutOf((await floats()).get('Columns'));
		const laid = lines(layout);
		expect(Math.max(...laid.map((line) => line.right))).toBeGreaterThan(layout.frame.widthPx / 2);
		expect(text(layout)).toBe(words('Columns', 40));
	});

	it('anchors the text in the middle of the box and keeps its formatting', async () => {
		const layout = layoutOf((await floats()).get('Formatted'));
		const [first, last] = [layout.blocks[0]!, layout.blocks.at(-1)!];
		const used = last.yPx + last.heightPx - first.yPx;
		expect(first.yPx).toBeCloseTo((layout.frame.heightPx - used) / 2, 6);
		expect(first.yPx).toBeGreaterThan(0);
		const heading = (layout.blocks[0] as LayoutParagraphBox).lines[0]!.fragments[0]!;
		expect(heading).toMatchObject({ bold: true, fontSizePt: 14 });
		// Right-aligned: the heading ends at the text rectangle's right edge.
		expect(lines(layout)[0]!.right).toBeCloseTo(layout.frame.widthPx, 0);
		expect((layout.blocks[1] as LayoutParagraphBox).frame?.borders?.bottom).toBeDefined();
	});

	it('keeps text clear of the preset’s text rectangle and half the outline, as Word does', async () => {
		const placed = await floats([
			{ ...box('Wrapped'), id: 1, name: 'Bare', outline: 0 },
			{ ...box('Wrapped'), id: 2, name: 'Thick', outline: 76200 },
			{ ...box('Wrapped'), id: 3, name: 'Ellipse', preset: 'ellipse', width: 2.5, height: 1.5 },
		]);
		const bare = layoutOf(placed.get('Bare'));
		const thick = layoutOf(placed.get('Thick'));
		expect(bare.frame.xPx).toBeCloseTo(9.6, 6);
		// A 6 pt (8 px) outline takes 4 px from each side.
		expect(thick.frame.xPx).toBeCloseTo(9.6 + 4, 6);
		expect(bare.frame.widthPx - thick.frame.widthPx).toBeCloseTo(8, 6);
		// An ellipse's text rectangle is inset by (1 - cos 45°) / 2 of its size on each side.
		const ellipse = layoutOf(placed.get('Ellipse'));
		const inset = (1 - Math.SQRT1_2) / 2;
		expect(ellipse.frame.xPx).toBeCloseTo(240 * inset + PAD.x, 1);
		expect(ellipse.frame.yPx).toBeCloseTo(144 * inset + PAD.y, 1);
	});

	it('draws a column break in a text box as a line break, as Word does', async () => {
		const placed = await floats([
			{
				...box('Wrapped'),
				content:
					'<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:p><w:r><w:br w:type="column"/><w:t>After</w:t></w:r></w:p>',
			},
		]);
		const laid = lines(layoutOf(placed.get('Wrapped')));
		expect(laid.map((line) => line.text.trim())).toEqual(['Before', '', 'After']);
	});

	it('saves an unchanged document byte for byte', async () => {
		const bytes = await textBoxLayoutFixture();
		const loaded = await loadDocx(bytes);
		const saved = await loaded.save(structuredClone(loaded.model));
		const xml = async (data: Uint8Array) =>
			(await JSZip.loadAsync(data)).file('word/document.xml')!.async('string');
		expect(await xml(saved)).toBe(textBoxDocumentXml(BOXES));
	});
});

function drawings(model: DocumentModel): InlineImage[] {
	return model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.flatMap((run) => (run.image ? [run.image] : [])) : [],
	);
}

/** Arial's advance widths (per 1000 em) for the characters of the fixture, from public font metrics. */
const ARIAL: Record<string, number> = {
	' ': 278,
	...Object.fromEntries([...'0123456789'].map((digit) => [digit, 556])),
	...Object.fromEntries([...'abdeghnopqu'].map((letter) => [letter, 556])),
	...Object.fromEntries([...'cksvxyz'].map((letter) => [letter, 500])),
	f: 278,
	t: 278,
	i: 222,
	j: 222,
	l: 222,
	m: 833,
	r: 333,
	w: 722,
	A: 667,
	B: 667,
	C: 722,
	F: 611,
	L: 556,
	S: 667,
	U: 722,
	V: 667,
	W: 944,
	':': 278,
	'.': 278,
};
/** Arial measured as Word 16 lays it out: its widths, and single lines of 1.15 em (12.65 pt at 11 pt). */
const arial: TextMeasurer = {
	widthOf: (text: string, font: LayoutFontSpec) =>
		[...text].reduce((sum, ch) => sum + ((ARIAL[ch] ?? 556) / 1000) * font.sizePx, 0),
	lineHeightOf: (font: LayoutFontSpec) => font.sizePx * 1.15,
	ascentOf: (font: LayoutFontSpec) => font.sizePx * 0.905,
};
const WORD_SAVED = new URL('./fixtures/text-boxes/word-saved.docx', import.meta.url);

describe('the Word-saved text boxes, against Word 16’s PDF of them', () => {
	async function wordFloats() {
		const bytes = new Uint8Array(await readFile(WORD_SAVED));
		const { model } = await loadDocx(bytes);
		const placed = layoutDocumentModel(model, arial).pages.flatMap((page) => page.floats ?? []);
		return new Map(BOXES.map((box, index) => [box.name, must(placed[index], box.name)]));
	}
	const first = (layout: TextBoxLayout, count: number) =>
		lines(layout)
			.slice(0, count)
			.map((line) => line.text.trim());

	it('breaks the lines where Word breaks them', async () => {
		const placed = await wordFloats();
		expect(first(layoutOf(placed.get('Wrapped')), 5)).toEqual([
			'Wrapped word1 word2',
			'word3 word4 word5',
			'word6 word7 word8',
			'word9 word10 word11',
			'word12 word13 word14',
		]);
		expect(first(layoutOf(placed.get('Columns')), 3)).toEqual([
			'Columns word1 word2 word3 word4',
			'word5 word6 word7 word8 word9',
			'word10 word11 word12 word13 word14',
		]);
		for (const name of ['Vertical', 'Vertical 270'])
			expect(first(layoutOf(placed.get(name)), 3).slice(1)).toEqual([
				'word4 word5 word6 word7 word8',
				'word9 word10 word11 word12',
			]);
		expect(first(layoutOf(placed.get('Autofit')), 6).map((line) => line.split(' ').at(-1))).toEqual(
			['word4', 'word9', 'word13', 'word17', 'word21', 'word24'],
		);
	});

	it('splits the linked chain after word8, as Word does', async () => {
		const placed = await wordFloats();
		expect(first(layoutOf(placed.get('Linked first')), 9)).toEqual([
			'Linked word1 word2',
			'word3 word4 word5',
			'word6 word7 word8',
		]);
		expect(first(layoutOf(placed.get('Linked second')), 2)).toEqual([
			'word9 word10 word11 word12',
			'word13 word14 word15 word16',
		]);
	});

	it('draws the autofit and unwrapped boxes at Word’s sizes', async () => {
		const placed = await wordFloats();
		const pt = (px: number) => (px * 72) / 96;
		// Word: 83.8 pt high, and 200.4 pt wide.
		expect(pt(layoutOf(placed.get('Autofit')).heightPx)).toBeCloseTo(83.8, 0);
		expect(pt(layoutOf(placed.get('No wrap')).widthPx)).toBeCloseTo(200.4, 0);
		expect(layoutOf(placed.get('Shrink')).heightPx).toBe(58);
	});

	it('saves Word’s file byte for byte', async () => {
		const bytes = new Uint8Array(await readFile(WORD_SAVED));
		const loaded = await loadDocx(bytes);
		const saved = await loaded.save(structuredClone(loaded.model));
		const xml = async (data: Uint8Array) =>
			(await JSZip.loadAsync(data)).file('word/document.xml')!.async('string');
		expect(await xml(saved)).toBe(await xml(bytes));
		expect(drawings(loaded.model)[7]?.textBoxContent?.link).toEqual({ id: 7, seq: 1 });
	});
});
