import { describe, expect, it } from 'vitest';
import type { TextRun } from '../index';
import { adaptDocumentModel } from './adapter';
import type { LayoutParagraph } from './input';

const page = {
	width: 816,
	height: 1056,
	marginTop: 96,
	marginRight: 96,
	marginBottom: 96,
	marginLeft: 96,
};
const symbol = { font: 'Wingdings', char: 'F0FE' };

function layoutRuns(runs: TextRun[]) {
	const input = adaptDocumentModel({
		blocks: [{ type: 'paragraph', id: 'p1', runs }],
		page,
		warnings: [],
	});
	return (input.sections[0]!.blocks[0] as LayoutParagraph).runs.map(
		({ text, fontFamily, sourceLength }) => ({ text, fontFamily, sourceLength }),
	);
}

describe('symbols in Print Layout', () => {
	it('draws a run made of its symbol in the symbol font', () => {
		expect(layoutRuns([{ text: '\uF0FE', symbol, fontFamily: 'Arial' }])).toEqual([
			{ text: '\uF0FE', fontFamily: 'Wingdings', sourceLength: 1 },
		]);
	});

	it('draws each symbol of a run that also holds text in the symbol font', () => {
		expect(layoutRuns([{ text: 'Done \uF0FE ok \uF0FE', symbol, fontFamily: 'Arial' }])).toEqual([
			{ text: 'Done ', fontFamily: 'Arial', sourceLength: 5 },
			{ text: '\uF0FE', fontFamily: 'Wingdings', sourceLength: 1 },
			{ text: ' ok ', fontFamily: 'Arial', sourceLength: 4 },
			{ text: '\uF0FE', fontFamily: 'Wingdings', sourceLength: 1 },
		]);
	});

	it('leaves a run without its symbol character whole', () => {
		expect(layoutRuns([{ text: 'plain', symbol, fontFamily: 'Arial' }])).toEqual([
			{ text: 'plain', fontFamily: 'Arial', sourceLength: 5 },
		]);
	});
});
