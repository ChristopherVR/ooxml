import { describe, expect, it } from 'vitest';
import {
	applyParagraphFormatPatch,
	applyRunFormatPatch,
	createDocument,
	findDocumentFormattedMatches,
	findFormattedMatches,
	hasFormatCriteria,
	hasReplaceFormat,
	matchesFormatCriteria,
	replaceDocumentFormattedMatches,
	replaceFormattedMatches,
	type Block,
	type DocumentModel,
	type Paragraph,
	type ParagraphStyleCatalog,
	type RunStyleCatalog,
	type Table,
	type TextRun,
	type ThemeCatalog,
	type Twips,
} from './index';

const para = (id: string, runs: TextRun[], extra: Partial<Paragraph> = {}): Paragraph => ({
	type: 'paragraph',
	id,
	runs,
	...extra,
});

const tableOf = (id: string, cellTexts: string[][]): Table => ({
	type: 'table',
	id,
	rows: cellTexts.map((row, r) =>
		row.map((text, c) => ({
			paragraphs: [para(`${id}-${r}-${c}`, [{ text, bold: r === 0 }])],
		})),
	),
});

const sampleStyles = (): {
	paragraphStyles: ParagraphStyleCatalog;
	runStyles: RunStyleCatalog;
	theme: ThemeCatalog;
} => ({
	paragraphStyles: {
		docDefaults: { align: 'left' },
		styles: {
			Normal: { id: 'Normal', name: 'Normal', isDefault: true, formatting: { align: 'left' } },
			Heading1: {
				id: 'Heading1',
				name: 'heading 1',
				basedOn: 'Normal',
				formatting: { align: 'center' },
			},
			Quote: {
				id: 'Quote',
				name: 'Quote',
				basedOn: 'Normal',
				formatting: { align: 'justify' },
			},
		},
		warnings: [],
	},
	runStyles: {
		docDefaults: { fontFamily: 'Calibri', fontSize: 11 },
		styles: {
			Heading1: {
				id: 'Heading1',
				type: 'paragraph',
				formatting: { bold: true, fontSize: 16, fontFamily: 'Cambria' },
			},
			Emphasis: {
				id: 'Emphasis',
				type: 'character',
				formatting: { italic: true, color: '#2F5496' },
			},
		},
		warnings: [],
	},
	theme: {
		colors: {
			dk1: '#000000',
			lt1: '#FFFFFF',
			dk2: '#1F497D',
			lt2: '#EEECE1',
			accent1: '#4F81BD',
			accent2: '#C0504D',
			accent3: '#9BBB59',
			accent4: '#8064A2',
			accent5: '#4BACC6',
			accent6: '#F79646',
			hlink: '#0000FF',
			folHlink: '#800080',
		},
		colorMapping: { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' },
		fonts: { major: { latin: 'Cambria' }, minor: { latin: 'Calibri' } },
	},
});

describe('hasFormatCriteria & hasReplaceFormat', () => {
	it('distinguishes empty criteria/patches from non-empty ones', () => {
		expect(hasFormatCriteria(undefined)).toBe(false);
		expect(hasFormatCriteria({})).toBe(false);
		expect(hasFormatCriteria({ bold: false })).toBe(true);
		expect(hasFormatCriteria({ styleId: 'Heading1' })).toBe(true);

		expect(hasReplaceFormat(undefined)).toBe(false);
		expect(hasReplaceFormat({})).toBe(false);
		expect(hasReplaceFormat({ color: null })).toBe(true);
		expect(hasReplaceFormat({ alignment: 'right' })).toBe(true);
	});
});

describe('matchesFormatCriteria: font, paragraph and style matching', () => {
	it('matches direct font properties (family, size, bold, italic, underline, color, highlight)', () => {
		const p = para('p1', []);
		const run: TextRun = {
			text: 'hello',
			fontFamily: 'Arial',
			fontSize: 12.5,
			bold: true,
			italic: false,
			underlineStyle: 'double',
			color: '#1a2b3c',
			highlight: 'yellow',
		};
		expect(
			matchesFormatCriteria(run, p, {
				fontFamily: '  arial ',
				fontSizePt: 12.5,
				bold: true,
				italic: false,
				underline: true,
				color: '1A2B3C',
				highlight: 'yellow',
			}),
		).toBe(true);

		expect(matchesFormatCriteria(run, p, { bold: false })).toBe(false);
		expect(matchesFormatCriteria(run, p, { italic: true })).toBe(false);
		expect(matchesFormatCriteria(run, p, { underline: false })).toBe(false);
		expect(matchesFormatCriteria(run, p, { fontSizePt: 12 })).toBe(false);
		expect(matchesFormatCriteria(run, p, { fontFamily: 'Georgia' })).toBe(false);
		expect(matchesFormatCriteria(run, p, { color: '#FFFFFF' })).toBe(false);
		expect(matchesFormatCriteria(run, p, { highlight: 'cyan' })).toBe(false);
		expect(matchesFormatCriteria(run, p, { highlight: 'none' })).toBe(false);
		expect(matchesFormatCriteria({ text: 'plain' }, p, { highlight: 'none' })).toBe(true);
		expect(
			matchesFormatCriteria({ text: 'u', underlineStyle: 'none' }, p, { underline: false }),
		).toBe(true);
	});

	it('resolves inherited style formatting and theme colors through context', () => {
		const context = sampleStyles();
		const heading = para('h1', [{ text: 'Title' }], { style: 'Heading1' });
		expect(
			matchesFormatCriteria(
				heading.runs[0]!,
				heading,
				{
					bold: true,
					fontSizePt: 16,
					fontFamily: 'Cambria',
					alignment: 'center',
					styleId: 'Heading1',
				},
				context,
			),
		).toBe(true);

		const normal = para('n1', [{ text: 'note', style: 'Emphasis' }]);
		expect(
			matchesFormatCriteria(
				normal.runs[0]!,
				normal,
				{ italic: true, color: '#2F5496', styleId: 'Emphasis', alignment: 'left' },
				context,
			),
		).toBe(true);
		expect(matchesFormatCriteria(normal.runs[0]!, normal, { styleId: 'Normal' }, context)).toBe(
			true,
		);

		const themedRun: TextRun = { text: 'accent', colorTheme: { token: 'accent1' } };
		expect(matchesFormatCriteria(themedRun, normal, { color: '#4F81BD' }, context)).toBe(true);
	});

	it('respects RTL default alignment when align is omitted', () => {
		const rtlPara = para('r1', [{ text: 'سلام' }], { direction: 'rtl' });
		expect(matchesFormatCriteria(rtlPara.runs[0]!, rtlPara, { alignment: 'right' })).toBe(true);
		expect(matchesFormatCriteria(rtlPara.runs[0]!, rtlPara, { alignment: 'left' })).toBe(false);
	});
});

describe('findFormattedMatches: text + format search', () => {
	it('matches only when every character of the match satisfies format criteria', () => {
		const blocks: Block[] = [
			para('p1', [
				{ text: 'foo', bold: true },
				{ text: ' ' },
				{ text: 'fo', bold: true },
				{ text: 'o' },
				{ text: ' ' },
				{ text: 'f', bold: true },
				{ text: 'oo', bold: true, italic: true },
			]),
		];
		const matches = findFormattedMatches(blocks, 'foo', { format: { bold: true } });
		expect(matches.map((m) => [m.start, m.end, m.text])).toEqual([
			[0, 3, 'foo'],
			[8, 11, 'foo'],
		]);
		expect(matches[1]).toMatchObject({
			startRunIndex: 5,
			startRunOffset: 0,
			endRunIndex: 6,
			endRunOffset: 2,
		});
	});

	it('combines format criteria with wildcards, matchCase and wholeWords', () => {
		const blocks: Block[] = [
			para('p1', [
				{ text: 'Cat cat catalog ', italic: true },
				{ text: 'cat', italic: false },
			]),
		];
		const exact = findFormattedMatches(blocks, 'cat', {
			matchCase: true,
			wholeWords: true,
			format: { italic: true },
		});
		expect(exact.map((m) => [m.start, m.text])).toEqual([[4, 'cat']]);

		const wild = findFormattedMatches(blocks, '<c*t>', {
			wildcards: true,
			format: { italic: true },
		});
		expect(wild.map((m) => m.text)).toEqual(['Cat', 'cat']);
	});

	it('searches inside tables without matching across cell boundaries', () => {
		const blocks: Block[] = [
			tableOf('t1', [
				['alpha', 'beta'],
				['alpha', 'gamma'],
			]),
		];
		const boldAlpha = findFormattedMatches(blocks, 'alpha', { format: { bold: true } });
		expect(boldAlpha).toHaveLength(1);
		expect(boldAlpha[0]?.paragraphId).toBe('t1-0-0');

		const crossCell = findFormattedMatches(blocks, 'alphabeta', { format: { bold: true } });
		expect(crossCell).toEqual([]);
	});
});

describe('findFormattedMatches: format-only search (empty query)', () => {
	it('returns empty array when both query and format criteria are empty', () => {
		const blocks: Block[] = [para('p1', [{ text: 'Hello', bold: true }])];
		expect(findFormattedMatches(blocks, '', {})).toEqual([]);
	});

	it('matches maximal contiguous spans within each paragraph without merging paragraphs', () => {
		const blocks: Block[] = [
			para('p1', [
				{ text: 'Plain ' },
				{ text: 'Bold ', bold: true },
				{ text: 'BoldItalic', bold: true, italic: true },
				{ text: ' end ' },
				{ text: 'Again', bold: true },
			]),
			para('p2', [{ text: 'NextParaBold', bold: true }]),
		];
		const matches = findFormattedMatches(blocks, '', { format: { bold: true } });
		expect(matches.map((m) => ({ id: m.paragraphId, text: m.text }))).toEqual([
			{ id: 'p1', text: 'Bold BoldItalic' },
			{ id: 'p1', text: 'Again' },
			{ id: 'p2', text: 'NextParaBold' },
		]);
	});

	it('matches whole paragraphs by paragraph style or alignment', () => {
		const context = sampleStyles();
		const blocks: Block[] = [
			para('p1', [{ text: 'Heading ' }, { text: 'One' }], { style: 'Heading1' }),
			para('p2', [{ text: 'Body text' }], { style: 'Normal' }),
			para('p3', [{ text: 'Centered body' }], { align: 'center' }),
		];
		const byStyle = findFormattedMatches(blocks, '', { format: { styleId: 'Heading1' }, context });
		expect(byStyle.map((m) => [m.paragraphId, m.text])).toEqual([['p1', 'Heading One']]);

		const byAlign = findFormattedMatches(blocks, '', { format: { alignment: 'center' }, context });
		expect(byAlign.map((m) => [m.paragraphId, m.text])).toEqual([
			['p1', 'Heading One'],
			['p3', 'Centered body'],
		]);
	});
});

describe('replaceFormattedMatches: format & text replacement', () => {
	it('applies run & paragraph format patches without losing untouched properties', () => {
		const originalRun: TextRun = {
			text: 'keep',
			bold: true,
			strike: true,
			language: 'en-US',
			fontSize: 12,
			color: '#111111',
		};
		const patchedRun = applyRunFormatPatch(originalRun, {
			italic: true,
			fontFamily: 'Georgia',
			color: '#cc0000',
			highlight: 'yellow',
		});
		expect(patchedRun).toEqual({
			text: 'keep',
			bold: true,
			italic: true,
			strike: true,
			language: 'en-US',
			fontFamily: 'Georgia',
			fontSize: 12,
			color: '#CC0000',
			highlight: 'yellow',
		});

		const originalPara = para('p1', [originalRun], {
			spacingBeforeTwips: 120 as Twips,
			justification: 'both',
			align: 'justify',
		});
		const patchedPara = applyParagraphFormatPatch(originalPara, {
			alignment: 'right',
			styleId: 'Quote',
		});
		expect(patchedPara.align).toBe('right');
		expect(patchedPara.justification).toBeUndefined();
		expect(patchedPara.style).toBe('Quote');
		expect(patchedPara.spacingBeforeTwips).toBe(120);
	});

	it('replaces format only when replacement text is empty and replaceFormat is set', () => {
		const blocks: Block[] = [
			para('p1', [
				{ text: 'Say ', fontSize: 11 },
				{ text: 'hello world', bold: true, strike: true, fontSize: 14 },
				{ text: '!' },
			]),
		];
		const res = replaceFormattedMatches(blocks, 'world', '', {
			format: { bold: true },
			replaceFormat: { italic: true, color: '#00AA00', underline: true },
		});
		expect(res.replacedCount).toBe(1);
		const p = res.blocks[0] as Paragraph;
		expect(p.runs.map((r) => r.text).join('')).toBe('Say hello world!');
		// Untouched prefix of the split run keeps original formatting
		expect(p.runs[1]).toEqual({ text: 'hello ', bold: true, strike: true, fontSize: 14 });
		// Matched slice keeps bold/strike/fontSize and gains italic/color/underline
		expect(p.runs[2]).toEqual({
			text: 'world',
			bold: true,
			italic: true,
			underline: true,
			strike: true,
			fontSize: 14,
			color: '#00AA00',
		});
		// Original input is never mutated
		expect((blocks[0] as Paragraph).runs).toHaveLength(3);
	});

	it('deletes matched text when replacement is empty and replaceFormat is omitted', () => {
		const blocks: Block[] = [para('p1', [{ text: 'drop me now', bold: true }])];
		const res = replaceFormattedMatches(blocks, 'me ', '', { format: { bold: true } });
		expect(res.replacedCount).toBe(1);
		expect((res.blocks[0] as Paragraph).runs.map((r) => r.text).join('')).toBe('drop now');
	});

	it('replaces text + formatting with Replace All and single Replace (matchIndex)', () => {
		const blocks: Block[] = [
			para('p1', [{ text: 'old and old', fontFamily: 'Arial', fontSize: 12 }], {
				align: 'left',
				spacingAfterTwips: 200 as Twips,
			}),
		];
		const single = replaceFormattedMatches(blocks, 'old', 'new', {
			matchIndex: 1,
			replaceFormat: { bold: true, highlight: 'cyan', alignment: 'center', styleId: 'Quote' },
		});
		expect(single.replacedCount).toBe(1);
		const pSingle = single.blocks[0] as Paragraph;
		expect(pSingle.runs.map((r) => r.text).join('')).toBe('old and new');
		expect(pSingle.runs[0]?.bold).toBeUndefined();
		expect(pSingle.runs[1]).toMatchObject({
			text: 'new',
			bold: true,
			highlight: 'cyan',
			fontFamily: 'Arial',
			fontSize: 12,
		});
		expect(pSingle.align).toBe('center');
		expect(pSingle.style).toBe('Quote');
		expect(pSingle.spacingAfterTwips).toBe(200);

		const all = replaceFormattedMatches(blocks, 'old', 'NEW', {
			replaceFormat: { bold: true, color: null, highlight: 'none' },
		});
		expect(all.replacedCount).toBe(2);
		expect((all.blocks[0] as Paragraph).runs.map((r) => r.text).join('')).toBe('NEW and NEW');
	});

	it('supports format-only find + format-only replace (e.g. convert all bold to italic + blue)', () => {
		const blocks: Block[] = [
			para('p1', [
				{ text: 'Normal ' },
				{ text: 'Strong', bold: true, underlineStyle: 'wave' },
				{ text: ' text' },
			]),
		];
		const res = replaceFormattedMatches(blocks, '', '', {
			format: { bold: true },
			replaceFormat: { bold: false, italic: true, underline: false, color: '#0000FF' },
		});
		expect(res.replacedCount).toBe(1);
		const p = res.blocks[0] as Paragraph;
		expect(p.runs.map((r) => r.text).join('')).toBe('Normal Strong text');
		expect(p.runs[1]).toEqual({
			text: 'Strong',
			bold: false,
			italic: true,
			underline: false,
			color: '#0000FF',
		});
	});

	it('applies character styles to runs and paragraph styles to paragraphs', () => {
		const context = sampleStyles();
		const blocks: Block[] = [para('p1', [{ text: 'Focus here' }], { style: 'Normal' })];
		const res = replaceFormattedMatches(blocks, 'Focus', '', {
			context,
			replaceFormat: { styleId: 'Emphasis' },
		});
		const p = res.blocks[0] as Paragraph;
		expect(p.style).toBe('Normal');
		expect(p.runs[0]?.style).toBe('Emphasis');
	});

	it('searches and replaces formatted matches across document stories', () => {
		const sharedHeader = { blocks: [para('h1', [{ text: 'Draft', bold: true }])] };
		const styles = sampleStyles();
		const section = (headers: { default: typeof sharedHeader }) =>
			({ ...createDocument().sections?.[0], headers }) as NonNullable<
				DocumentModel['sections']
			>[number];
		const model: DocumentModel = {
			...createDocument(),
			paragraphStyles: styles.paragraphStyles,
			characterStyles: styles.runStyles,
			theme: styles.theme,
			blocks: [para('b1', [{ text: 'Draft', bold: true }, { text: ' plain Draft' }])],
			sections: [section({ default: sharedHeader }), section({ default: sharedHeader })],
			footnotes: [{ id: '1', blocks: [para('f1', [{ text: 'Draft', bold: true }])] }],
		};

		const found = findDocumentFormattedMatches(model, 'Draft', { format: { bold: true } });
		expect(found.map((m) => m.storyKey)).toEqual(['main', 'header:0:default', 'footnote:1']);

		const replacedAll = replaceDocumentFormattedMatches(model, 'Draft', 'Final', {
			format: { bold: true },
			replaceFormat: { italic: true, color: '#123456' },
		});
		expect(replacedAll.replacedCount).toBe(3);
		expect((replacedAll.model.blocks[0] as Paragraph).runs[0]).toMatchObject({
			text: 'Final',
			bold: true,
			italic: true,
			color: '#123456',
		});
		expect((replacedAll.model.blocks[0] as Paragraph).runs[1]?.text).toBe(' plain Draft');
		const allHeader = replacedAll.model.sections?.[0]?.headers?.default?.blocks[0] as
			| Paragraph
			| undefined;
		expect(allHeader?.runs[0]?.text).toBe('Final');
		const allNote = replacedAll.model.footnotes?.[0]?.blocks[0] as Paragraph | undefined;
		expect(allNote?.runs[0]?.text).toBe('Final');

		// Single matchIndex=1 targets only the header story
		const replacedOne = replaceDocumentFormattedMatches(model, 'Draft', 'HeaderOnly', {
			format: { bold: true },
			matchIndex: 1,
		});
		expect(replacedOne.replacedCount).toBe(1);
		expect((replacedOne.model.blocks[0] as Paragraph).runs[0]?.text).toBe('Draft');
		const oneHeader = replacedOne.model.sections?.[0]?.headers?.default?.blocks[0] as
			| Paragraph
			| undefined;
		expect(oneHeader?.runs[0]?.text).toBe('HeaderOnly');
		const oneNote = replacedOne.model.footnotes?.[0]?.blocks[0] as Paragraph | undefined;
		expect(oneNote?.runs[0]?.text).toBe('Draft');
	});
});
