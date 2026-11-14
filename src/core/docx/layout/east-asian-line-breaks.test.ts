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
		const result = lines([{ text: 'aaa bbb' }, { text: 'bb', bold: true }], {}, 80);
		expect(result).toEqual(['aaa ', 'bbbbb'].map((line) => line.trimEnd()));
		expect(at(result, 1)).toBe('bbbbb');
	});
});
