import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutParagraph, LayoutRun } from './input';
import type { TextMeasurer } from './measure';
import { at } from './test-helpers';

// Fixed pitch: every character is 10px wide, as in Word's probes (10 pt MS Mincho, 100 pt line).
const measurer: TextMeasurer = { widthOf: (text) => [...text].length * 10, lineHeightOf: () => 20 };
const noop = () => {};
const ja = (text: string, extra: Partial<LayoutRun> = {}): LayoutRun => ({
	text,
	kinsokuLanguage: 'ja',
	...extra,
});
const lines = (runs: LayoutRun[], extra: Partial<LayoutParagraph> = {}, width = 100) =>
	layoutParagraph({ kind: 'paragraph', id: 'p', runs, ...extra }, width, measurer, noop).lines.map(
		(line) => line.fragments.map((fragment) => fragment.text).join(''),
	);
const ten = '一'.repeat(10);

describe('kinsoku line breaking (Word 16 probes)', () => {
	it('moves the character before a prohibited line-start mark down with it', () => {
		expect(lines([ja(`${ten}。一一`)])).toEqual(['一'.repeat(9), '一。一一']);
	});

	it('keeps a prohibited line-end bracket with the character after it', () => {
		expect(lines([ja(`${'一'.repeat(9)}（一一`)])).toEqual(['一'.repeat(9), '（一一']);
	});

	it('applies across runs, by the language of the run the mark is in', () => {
		expect(lines([ja(ten), ja('」一', { bold: true })])).toEqual(['一'.repeat(9), '一」一']);
		expect(lines([ja(ten), { text: '」一', bold: true }])).toEqual([ten, '」一']);
	});

	it('does nothing when the paragraph turns kinsoku off or the text has no East Asian language', () => {
		expect(lines([ja(`${ten}。一`)], { kinsoku: false })).toEqual([ten, '。一']);
		expect(lines([{ text: `${ten}。一` }])).toEqual([ten, '。一']);
	});

	it('uses the custom sets of the document for their language', () => {
		const kinsokuSettings = { noLineBreaksBefore: { language: 'ja-JP', chars: '々' } };
		expect(lines([ja(`${ten}。一`)], { kinsokuSettings })).toEqual([ten, '。一']);
		expect(lines([ja(`${ten}々一`)], { kinsokuSettings })).toEqual(['一'.repeat(9), '一々一']);
		expect(
			lines([{ text: `${ten}。一`, kinsokuLanguage: 'zh-Hans' }], { kinsokuSettings }),
		).toEqual(['一'.repeat(9), '一。一']);
	});

	it('uses the strict Japanese set for small kana when the document asks for it', () => {
		expect(lines([ja(`${ten}ッ一`)])).toEqual([ten, 'ッ一']);
		expect(lines([ja(`${ten}ッ一`)], { kinsokuSettings: { strict: true } })).toEqual([
			'一'.repeat(9),
			'一ッ一',
		]);
	});

	it('breaks anyway when the glued chain fills the whole line', () => {
		expect(lines([ja('一。。。。')], {}, 30)).toEqual(['一。。', '。。']);
	});
});

describe('words split across runs', () => {
	it('never breaks inside a word whose letters are formatted differently', () => {
		expect(lines([{ text: 'aaa bbb' }, { text: 'bb', bold: true }], {}, 80)).toEqual([
			'aaa',
			'bbbbb',
		]);
	});
});

describe('automatic East Asian and Western spacing (w:autoSpaceDE, w:autoSpaceDN)', () => {
	// 7.5 pt is 10px, so a quarter em is 2.5px.
	const run = (text: string, fontSizePt = 7.5): LayoutRun => ({ text, fontSizePt });
	const fragmentsOf = (runs: LayoutRun[], extra: Partial<LayoutParagraph> = {}) =>
		at(
			layoutParagraph({ kind: 'paragraph', id: 'p', runs, ...extra }, 400, measurer, noop).lines,
			0,
		).fragments;
	const positions = (runs: LayoutRun[], extra: Partial<LayoutParagraph> = {}) =>
		fragmentsOf(runs, extra).map((fragment) => [fragment.text, fragment.xPx]);

	it('adds a quarter em after the character before each East Asian/Western boundary', () => {
		expect(positions([run('日本Wa日本12日x')])).toEqual([
			['日', 0],
			['本', 10],
			['Wa', 22.5],
			['日', 45],
			['本', 55],
			['12', 67.5],
			['日', 90],
			['x', 102.5],
		]);
	});

	it('spaces across runs by the preceding size and leaves the glyph widths alone', () => {
		const fragments = fragmentsOf([run('日'), run('a', 15), run('日')]);
		expect(fragments.map((fragment) => [fragment.xPx, fragment.widthPx])).toEqual([
			[0, 10],
			[12.5, 10],
			[27.5, 10],
		]);
	});

	it('turns letters and digits off separately', () => {
		expect(positions([run('日本Wa日本12日x')], { autoSpaceDE: false })).toEqual([
			['日', 0],
			['本', 10],
			['Wa', 20],
			['日', 40],
			['本', 50],
			['12', 62.5],
			['日', 85],
			['x', 95],
		]);
		expect(positions([run('本12日')], { autoSpaceDN: false }).map(([, x]) => x)).toEqual([
			0, 10, 30,
		]);
	});

	it('does not space punctuation or text already separated by a space', () => {
		expect(positions([run('日 a.日(a)')]).map(([, x]) => x)).toEqual([0, 10, 20, 40, 50]);
	});

	it('does not let the gap after a line-end character push it to the next line', () => {
		expect(lines([run(`${ten}a`)])).toEqual([ten, 'a']);
	});
});

describe('Thai, Lao, Khmer and Myanmar line breaking', () => {
	const samples = {
		thai: 'ภาษาไทยเป็นภาษาที่มีระดับเสียงของคำแน่นอนหรือวรรณยุกต์เช่นเดียวกับภาษาจีน',
		lao: 'ພາສາລາວເປັນພາສາທີ່ມີສຽງວັນນະຍຸດແລະເປັນພາສາລາຊະການຂອງປະເທດລາວ',
		khmer: 'ភាសាខ្មែរគឺជាភាសាកំណើតរបស់ជនជាតិខ្មែរនិងជាភាសាផ្លូវការរបស់ប្រទេសកម្ពុជា',
		myanmar: 'မြန်မာဘာသာစကားသည်မြန်မာနိုင်ငံ၏ရုံးသုံးဘာသာစကားဖြစ်သည်',
	};
	for (const [script, text] of Object.entries(samples))
		it(`wraps a long unspaced ${script} run at grapheme cluster boundaries`, () => {
			const notes: string[] = [];
			const result = layoutParagraph(
				{ kind: 'paragraph', id: 'p', runs: [{ text }] },
				100,
				measurer,
				(note) => notes.push(note),
			);
			const texts = result.lines.map((line) => line.fragments.map((f) => f.text).join(''));
			expect(texts.length).toBeGreaterThan(3);
			expect(texts.join('')).toBe(text);
			expect(notes).toEqual([]);
			const boundaries = new Set(
				[...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text)].map(
					(part) => part.index,
				),
			);
			let offset = 0;
			for (const line of texts) {
				expect(boundaries.has(offset), `${script} line at ${offset}`).toBe(true);
				expect(measurer.widthOf(line, { family: '', sizePx: 10 })).toBeLessThanOrEqual(100);
				offset += line.length;
			}
		});

	it('does not break between Thai and an adjacent Latin word', () => {
		expect(lines([{ text: 'a กabcd' }], {}, 50)).toEqual(['a', 'กabcd']);
	});

	it('breaks after a zero width space', () => {
		expect(lines([{ text: 'aaaaaaaa​bbbbbbbb' }])).toEqual(['aaaaaaaa​', 'bbbbbbbb']);
	});
});

describe('hanging punctuation (w:overflowPunct)', () => {
	const justify = { align: 'justify' as const };
	const nine = '一'.repeat(9);

	it('hangs a line-end mark past the margin in a justified paragraph', () => {
		const result = layoutParagraph(
			{ kind: 'paragraph', id: 'p', runs: [ja(`${ten}。一一`)], ...justify },
			100,
			measurer,
			noop,
		);
		const first = at(result.lines, 0).fragments;
		expect(first.map((fragment) => fragment.text).join('')).toBe(`${ten}。`);
		expect(first.at(-1)?.xPx).toBe(100);
		expect(lines([ja(ten), ja('、一', { bold: true })], justify)).toEqual([`${ten}、`, '一']);
	});

	it('wraps instead when the paragraph is not justified or turns it off', () => {
		expect(lines([ja(`${ten}。一`)])).toEqual([nine, '一。一']);
		expect(lines([ja(`${ten}。一`)], { ...justify, overflowPunct: false })).toEqual([
			nine,
			'一。一',
		]);
	});

	it('hangs only marks of the hanging set, and only under kinsoku', () => {
		expect(lines([ja(`${ten}！一`)], justify)).toEqual([nine, '一！一']);
		expect(lines([ja(`${ten}。一`)], { ...justify, kinsoku: false })).toEqual([ten, '。一']);
	});

	it('does not hang a mark followed by another mark that may not start a line', () => {
		expect(lines([ja(`${ten}。」一`)], justify)).toEqual([nine, '一。」一']);
		expect(lines([ja(`${nine}。」一`)], justify)).toEqual([`${nine}。」`, '一']);
	});
});
