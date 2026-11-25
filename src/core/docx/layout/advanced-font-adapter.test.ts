import { createDocument, halfPoints } from '../index';
import { describe, expect, it } from 'vitest';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import { at } from './test-helpers';

describe('advanced font layout adaptation', () => {
	it('keeps explicit baseline full size even with inherited superscript', () => {
		const model = createDocument();
		model.characterStyles = {
			docDefaults: { verticalAlign: 'superscript', fontSize: 12 },
			styles: {},
			warnings: [],
		};
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [{ text: 'Script' }, { text: 'Plain', verticalAlign: 'baseline' }],
			},
		];
		const paragraph = adaptDocumentModel(model).sections[0]!.blocks[0]!;
		if (paragraph.kind !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs[0]!.script).toBe('super');
		expect(paragraph.runs[1]!.script).toBeUndefined();
		expect(paragraph.runs[1]!.fontSizePt).toBe(12);
	});
	it('adapts direct and inherited settings to layout units without flattening the model', () => {
		const model = createDocument();
		model.characterStyles = {
			docDefaults: { textScalePercent: 125 },
			styles: {},
			warnings: [],
		};
		model.blocks = [
			{
				type: 'paragraph',
				id: 'advanced',
				runs: [
					{ text: 'Inherited scale' },
					{ text: 'Position', positionHalfPoints: halfPoints(-6) },
					{ text: 'Kerning', kerningHalfPoints: halfPoints(24) },
				],
			},
		];
		const notes: string[] = [];
		const input = adaptDocumentModel(model, (note) => notes.push(note));
		expect(notes).toEqual([]);
		const paragraph = input.sections[0]!.blocks[0]!;
		if (paragraph.kind !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs[0]).toMatchObject({ textScalePercent: 125 });
		expect(paragraph.runs[1]).toMatchObject({ textScalePercent: 125, positionPx: -4 });
		expect(paragraph.runs[2]).toMatchObject({ textScalePercent: 125, kerningThresholdPt: 12 });
		expect(model.blocks[0]).toMatchObject({
			runs: [{ text: 'Inherited scale' }, { positionHalfPoints: -6 }, { kerningHalfPoints: 24 }],
		});
	});

	it('does not warn about neutral scale and baseline position', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'neutral',
				runs: [{ text: 'Normal', textScalePercent: 100, positionHalfPoints: halfPoints(0) }],
			},
		];
		const notes: string[] = [];
		adaptDocumentModel(model, (note) => notes.push(note));
		expect(notes).toEqual([]);
	});
});

it('wraps a Word paragraph without kerning unless its run formatting enables it', () => {
	const model = createDocument();
	model.page = { ...model.page, width: 100, marginLeft: 0, marginRight: 0 };
	model.blocks = [undefined, halfPoints(24)].map((kern, index) => ({
		type: 'paragraph',
		id: `p${index}`,
		runs: [
			{ text: 'AV AV', fontSize: 12, ...(kern === undefined ? {} : { kerningHalfPoints: kern }) },
		],
	}));
	const measurer = {
		widthOf: (text: string, font: { kerning?: string }) =>
			text === 'AV' ? (font.kerning === 'none' ? 52 : 40) : 4,
		lineHeightOf: () => 20,
	};
	const { pages } = layoutDocumentModel(model, measurer);
	const blocks = at(at(pages, 0).columns, 0).blocks;
	expect(blocks.map((block) => (block.kind === 'paragraph' ? block.lines.length : 0))).toEqual([
		2, 1,
	]);
});
