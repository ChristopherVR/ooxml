import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph, type TextRun } from '../index';
import { collectDocxWebfontRuns } from './webfont-runs';

const paragraph = (id: string, runs: TextRun[]): Paragraph => ({ type: 'paragraph', id, runs });

function modelWithStyles(): DocumentModel {
	const model = createDocument();
	model.theme = {
		colors: {},
		colorMapping: {},
		fonts: { major: { latin: 'Aptos Display' }, minor: { latin: 'Aptos', eastAsia: 'MS Mincho' } },
	};
	model.characterStyles = {
		docDefaults: { fontTheme: { ascii: 'minor', hAnsi: 'minor', eastAsia: 'minor' } },
		styles: {
			Heading1: {
				id: 'Heading1',
				type: 'paragraph',
				formatting: { fontTheme: { ascii: 'major', hAnsi: 'major' }, bold: true },
			},
			Code: { id: 'Code', type: 'character', formatting: { fontFamily: 'Courier New' } },
		},
		warnings: [],
	};
	model.paragraphStyles = {
		docDefaults: {},
		styles: { Heading1: { id: 'Heading1', formatting: {} } },
		warnings: [],
	};
	return model;
}

describe('collectDocxWebfontRuns', () => {
	it('resolves each run to the family it is drawn in: theme, paragraph and character styles', () => {
		const model = modelWithStyles();
		model.blocks = [
			{ type: 'paragraph', id: 'h', style: 'Heading1', runs: [{ text: 'Title' }] },
			{
				type: 'paragraph',
				id: 'b',
				runs: [
					{ text: 'Body ' },
					{ text: 'code()', style: 'Code' },
					{ text: 'Direct', fontFamily: 'Georgia', italic: true },
				],
			},
		];
		expect(collectDocxWebfontRuns(model)).toEqual([
			{ family: 'Aptos Display', bold: true, italic: false, text: 'Title' },
			{ family: 'Aptos', bold: false, italic: false, text: 'Body ' },
			{ family: 'Courier New', bold: false, italic: false, text: 'code()' },
			{ family: 'Georgia', bold: false, italic: true, text: 'Direct' },
			{ family: 'Aptos', bold: false, italic: false, text: '' },
		]);
	});

	it('splits east Asian text into its own font, as layout does', () => {
		const model = modelWithStyles();
		model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hi 日本' }] }];
		const families = collectDocxWebfontRuns(model).map((run) => run.family);
		expect(families).toContain('Aptos');
		expect(families).toContain('MS Mincho');
	});

	it('defaults an unstyled run to Calibri and skips hidden text', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [{ text: 'Plain' }, { text: 'Hidden', fontFamily: 'Cambria', vanish: true }],
			},
		];
		expect(collectDocxWebfontRuns(model)).toEqual([
			{ family: 'Calibri', bold: false, italic: false, text: 'Plain' },
			{ family: 'Calibri', bold: false, italic: false, text: '' },
		]);
	});

	it('covers tables (with their style), headers, footers and notes', () => {
		const model = createDocument();
		model.tableStyles = {
			styles: {
				Grid: { id: 'Grid', conditional: { wholeTable: { run: { fontFamily: 'Arial' } } } },
			},
			warnings: [],
		};
		model.blocks = [
			{
				type: 'table',
				id: 't',
				style: 'Grid',
				rows: [[{ paragraphs: [paragraph('c', [{ text: 'Cell' }])] }]],
			},
		];
		model.sections = [
			{
				headers: {
					default: { blocks: [paragraph('hd', [{ text: 'Head', fontFamily: 'Tinos' }])] },
				},
				footers: {
					first: { blocks: [paragraph('ft', [{ text: 'Foot', fontFamily: 'Cambria' }])] },
				},
			} as unknown as NonNullable<DocumentModel['sections']>[number],
		];
		model.footnotes = [
			{ id: '1', blocks: [paragraph('fn', [{ text: 'Note', fontFamily: 'Times New Roman' }])] },
		];
		const families = collectDocxWebfontRuns(model).map((run) => `${run.family}:${run.text}`);
		expect(families).toEqual(
			expect.arrayContaining(['Tinos:Head', 'Cambria:Foot', 'Times New Roman:Note']),
		);
		expect(families).toContain('Arial:Cell');
	});

	it('keeps the font of an empty styled paragraph, which the caret is drawn in', () => {
		const model = modelWithStyles();
		model.blocks = [{ type: 'paragraph', id: 'h', style: 'Heading1', runs: [] }];
		expect(collectDocxWebfontRuns(model)[0]).toEqual({
			family: 'Aptos Display',
			bold: true,
			italic: false,
			text: '',
		});
	});

	it('draws a symbol in its own font', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [{ text: '', symbol: { font: 'Wingdings', char: 'F0B7' } }],
			},
		];
		expect(collectDocxWebfontRuns(model)[0]?.family).toBe('Wingdings');
	});
});
