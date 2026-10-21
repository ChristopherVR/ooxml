import { describe, expect, it } from 'vitest';
import type { DocumentModel, FormField, TextRun } from '../index';
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
const checkbox: FormField = { kind: 'checkbox', checkBox: { default: false, checked: true } };
const dropDown: FormField = { kind: 'dropDown', dropDown: { entries: ['New', 'Revised'] } };
const field = (formField: FormField, instr: string, result?: string): TextRun[] => [
	{ text: '', fieldChar: 'begin', formField },
	{ text: '', fieldCode: ` ${instr} ` },
	...(result === undefined
		? []
		: [{ text: '', fieldChar: 'separate' } as TextRun, { text: result, field: { instr } }]),
	{ text: '', fieldChar: 'end' },
];

function layoutRuns(model: Partial<DocumentModel>) {
	const input = adaptDocumentModel({
		blocks: [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					...field(checkbox, 'FORMCHECKBOX'),
					...field(dropDown, 'FORMDROPDOWN'),
					...field({ kind: 'text' }, 'FORMTEXT', 'Ada'),
				],
			},
		],
		page,
		warnings: [],
		...model,
	});
	return (input.sections[0]!.blocks[0] as LayoutParagraph).runs.filter((run) => run.text);
}

describe('legacy form fields in Print Layout', () => {
	it('draws check boxes and drop-downs from their data and shades every form field', () => {
		expect(layoutRuns({}).map(({ text, background }) => ({ text, background }))).toEqual([
			{ text: '☒', background: '#d9d9d9' },
			{ text: 'New', background: '#d9d9d9' },
			{ text: 'Ada', background: '#d9d9d9' },
		]);
	});

	it('leaves form fields unshaded when the document turns shading off', () => {
		expect(layoutRuns({ formFieldShading: false }).every((run) => !run.background)).toBe(true);
	});
});
