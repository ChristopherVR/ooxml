import { describe, expect, it } from 'vitest';
import {
	describeParagraphFormatChange,
	describeRunFormatChange,
	formatChangeText,
	wordColorName,
} from './format-change-description';
import { loadDocx } from './parse';
import type { Paragraph, TextRun } from './model';
import { saveDocx } from './save';
import { signedTwips, twips } from './units';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const rPr = (inner: string) => `<w:rPr xmlns:w="${W}">${inner}</w:rPr>`;
const pPr = (inner: string) => `<w:pPr xmlns:w="${W}">${inner}</w:pPr>`;
const changed = (run: Omit<TextRun, 'text'>, previous: string): TextRun => ({
	text: 'Text',
	...run,
	revision: {
		id: '1',
		kind: 'formatChange',
		author: 'Ada',
		previousRunPropertiesXml: rPr(previous),
	},
});

describe('describeRunFormatChange', () => {
	it('words font changes as Word balloons do', () => {
		expect(describeRunFormatChange(changed({ bold: true }, ''))).toEqual(['Font: Bold']);
		expect(describeRunFormatChange(changed({}, '<w:b/><w:i/>'))).toEqual([
			'Font: Not Bold, Not Italic',
		]);
		expect(
			describeRunFormatChange(
				changed({ fontFamily: 'Arial', fontSize: 14, color: '#FF0000', underline: true }, ''),
			),
		).toEqual(['Font: Arial, 14 pt, Underline', 'Font color: Red']);
		expect(
			describeRunFormatChange(
				changed({ characterSpacingTwips: signedTwips(20), highlight: 'yellow' }, ''),
			),
		).toEqual(['Highlight', 'Expanded by 1 pt']);
	});

	it('reads a format change nested in an insertion, and gives up without a snapshot', () => {
		const run: TextRun = {
			text: 'x',
			italic: true,
			revision: { id: '1', kind: 'insert', author: 'Ada' },
			formatRevision: {
				id: '2',
				kind: 'formatChange',
				author: 'Ada',
				previousRunPropertiesXml: rPr(''),
			},
		};
		expect(describeRunFormatChange(run)).toEqual(['Font: Italic']);
		expect(describeRunFormatChange({ text: 'x' })).toBeUndefined();
		expect(
			describeRunFormatChange({
				text: 'x',
				revision: { id: '1', kind: 'formatChange', author: 'Ada' },
			}),
		).toBeUndefined();
	});
});

describe('describeParagraphFormatChange', () => {
	const paragraph = (properties: Partial<Paragraph>, previous: string): Paragraph => ({
		type: 'paragraph',
		id: 'p',
		runs: [{ text: 'Text' }],
		...properties,
		formatRevision: {
			id: '3',
			kind: 'paragraphChange',
			author: 'Ada',
			previousParagraphPropertiesXml: pPr(previous),
		},
	});

	it('words alignment, indents, spacing and styles', () => {
		expect(describeParagraphFormatChange(paragraph({ align: 'center' }, ''))).toEqual(['Centered']);
		expect(
			describeParagraphFormatChange(
				paragraph(
					{
						indentLeftTwips: signedTwips(720),
						hangingTwips: twips(360),
						spacingAfterTwips: twips(120),
					},
					'<w:spacing w:after="0"/>',
				),
			),
		).toEqual(['Indent: Left: 0.5", Hanging: 0.25"', 'Space After: 6 pt']);
		expect(
			describeParagraphFormatChange(
				paragraph({ lineSpacingTwips: signedTwips(360), lineSpacingRule: 'auto' }, ''),
			),
		).toEqual(['Line spacing: 1.5 lines']);
		const styled = paragraph({ style: 'Heading1' }, '');
		expect(
			describeParagraphFormatChange(styled, {
				paragraphStyles: {
					docDefaults: {},
					styles: { Heading1: { id: 'Heading1', name: 'heading 1', formatting: {} } },
					warnings: [],
				},
			}),
		).toEqual(['heading 1']);
		expect(describeParagraphFormatChange({ type: 'paragraph', id: 'x', runs: [] })).toBeUndefined();
	});

	it('describes the changes of a loaded document', async () => {
		const xml = `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Ada"><w:pPr/></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><w:color w:val="00B050"/><w:rPrChange w:id="2" w:author="Ada"><w:rPr/></w:rPrChange></w:rPr><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;
		const JSZip = (await import('jszip')).default;
		const zip = new JSZip();
		zip.file('word/document.xml', xml);
		const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const reloaded = (await loadDocx(await saveDocx(model))).model;
		const first = reloaded.blocks[0] as Paragraph;
		expect(formatChangeText(describeParagraphFormatChange(first)!)).toBe('Formatted: Centered');
		expect(formatChangeText(describeRunFormatChange(first.runs[0]!)!)).toBe(
			'Formatted: Font: Bold, Font color: Green',
		);
		expect(formatChangeText([])).toBe('Formatted');
	});
});

describe('wordColorName', () => {
	it('names standard, theme and custom colours', () => {
		expect(wordColorName('#c00000')).toBe('Dark Red');
		expect(wordColorName('123456')).toBe('Custom Color(RGB(18,52,86))');
		expect(wordColorName(undefined)).toBe('Auto');
		expect(wordColorName(undefined, { token: 'accent1' })).toBe('Accent 1');
		expect(wordColorName(undefined, { token: 'accent2', tint: 0.6 })).toBe('Accent 2, Lighter 40%');
		expect(wordColorName(undefined, { token: 'text1', shade: 0.75 })).toBe('Text 1, Darker 25%');
	});
});
