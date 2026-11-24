import type { TablePptxElement, XmlObject } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	tableCellsChangeCasePatch,
	tableCellsTextFormat,
	tableCellsTextStylePatch,
	tableTextShadowUpdate,
} from './table-cells-format';

function sampleTable(): TablePptxElement {
	return {
		id: 'tbl-1',
		type: 'table',
		x: 0,
		y: 0,
		width: 400,
		height: 120,
		tableData: {
			columnWidths: [0.5, 0.5],
			rows: [
				{
					cells: [
						{
							text: 'hello world',
							textRuns: [{ text: 'hello ' }, { text: 'world', bold: true }],
							style: { bold: true },
						},
						{
							text: 'Second cell',
							style: {},
						},
					],
				},
			],
		},
		rawXml: {
			'a:graphic': {
				'a:graphicData': {
					'a:tbl': {
						'a:tr': [
							{
								'a:tc': [
									{
										'a:txBody': {
											'a:p': {
												'a:r': [
													{ 'a:rPr': {}, 'a:t': 'hello ' },
													{ 'a:rPr': { '@_b': '1' }, 'a:t': 'world' },
												],
											},
										},
									},
									{
										'a:txBody': {
											'a:p': {
												'a:r': { 'a:rPr': {}, 'a:t': 'Second cell' },
											},
										},
									},
								],
							},
						],
					},
				},
			},
		},
	} as TablePptxElement;
}

describe('tableCellsTextStylePatch (highlight, spacing, shadow)', () => {
	it('applies highlightColor, characterSpacing and textShadow to selected cells and rawXml', () => {
		const table = sampleTable();
		const patch = tableCellsTextStylePatch(table, [{ row: 0, col: 0 }], {
			highlightColor: '#ffff00',
			characterSpacing: 150,
			...tableTextShadowUpdate(false),
		});
		const next = { ...table, ...patch } as TablePptxElement;
		const cell0 = next.tableData!.rows[0]!.cells[0]!;
		const cell1 = next.tableData!.rows[0]!.cells[1]!;

		expect(cell0.style?.textShadowColor).toBe('#000000');
		expect(cell0.textRuns).toMatchObject([
			{
				text: 'hello ',
				highlightColor: '#ffff00',
				characterSpacing: 150,
				textShadowColor: '#000000',
			},
			{
				text: 'world',
				bold: true,
				highlightColor: '#ffff00',
				characterSpacing: 150,
				textShadowColor: '#000000',
			},
		]);
		expect(cell1.style?.textShadowColor).toBeUndefined();

		const fmt = tableCellsTextFormat(next, [{ row: 0, col: 0 }]);
		expect(fmt.highlightColor).toBe('#ffff00');
		expect(fmt.characterSpacing).toBe(150);
		expect(fmt.textShadowColor).toBe('#000000');

		const r0 = (
			(
				(((next.rawXml as XmlObject)['a:graphic'] as XmlObject)['a:graphicData'] as XmlObject)[
					'a:tbl'
				] as XmlObject
			)['a:tr'] as XmlObject[]
		)[0]!['a:tc'] as XmlObject[];
		const runs = ((r0[0]!['a:txBody'] as XmlObject)['a:p'] as XmlObject)['a:r'] as XmlObject[];
		expect((runs[0]!['a:rPr'] as XmlObject)['@_spc']).toBe('150');
		expect((runs[0]!['a:rPr'] as XmlObject)['a:highlight']).toStrictEqual({
			'a:srgbClr': { '@_val': 'FFFF00' },
		});
		expect(
			((runs[0]!['a:rPr'] as XmlObject)['a:effectLst'] as XmlObject)['a:outerShdw'],
		).toBeDefined();

		// Toggling shadow off clears it from cell style, runs and rawXml.
		const offPatch = tableCellsTextStylePatch(
			next,
			[{ row: 0, col: 0 }],
			tableTextShadowUpdate(true),
		);
		const afterOff = { ...next, ...offPatch } as TablePptxElement;
		expect(afterOff.tableData!.rows[0]!.cells[0]!.style?.textShadowColor).toBeUndefined();
		expect(tableCellsTextFormat(afterOff, [{ row: 0, col: 0 }]).textShadowColor).toBeUndefined();
	});

	it('materializes textRuns on a plain-text cell so highlight and spacing paint immediately', () => {
		const table = sampleTable();
		const patch = tableCellsTextStylePatch(table, [{ row: 0, col: 1 }], {
			highlightColor: '#00ff00',
			characterSpacing: -75,
		});
		const cell1 = patch.tableData!.rows[0]!.cells[1]!;
		expect(cell1.textRuns).toMatchObject([
			{ text: 'Second cell', highlightColor: '#00ff00', characterSpacing: -75 },
		]);
	});
});

describe('tableCellsChangeCasePatch', () => {
	it('rewrites cell text, textRuns and rawXml for selected cells only across case modes', () => {
		const table = sampleTable();
		const upper = tableCellsChangeCasePatch(table, [{ row: 0, col: 0 }], 'upper');
		const next = { ...table, ...upper } as TablePptxElement;
		expect(next.tableData!.rows[0]!.cells[0]!.text).toBe('HELLO WORLD');
		expect(next.tableData!.rows[0]!.cells[0]!.textRuns).toStrictEqual([
			{ text: 'HELLO ' },
			{ text: 'WORLD', bold: true },
		]);
		expect(next.tableData!.rows[0]!.cells[1]!.text).toBe('Second cell');

		const r0 = (
			(
				(((next.rawXml as XmlObject)['a:graphic'] as XmlObject)['a:graphicData'] as XmlObject)[
					'a:tbl'
				] as XmlObject
			)['a:tr'] as XmlObject[]
		)[0]!['a:tc'] as XmlObject[];
		const runs = ((r0[0]!['a:txBody'] as XmlObject)['a:p'] as XmlObject)['a:r'] as XmlObject[];
		expect(runs[0]!['a:t']).toBe('HELLO ');
		expect(runs[1]!['a:t']).toBe('WORLD');

		const lower = tableCellsChangeCasePatch(next, [{ row: 0, col: 0 }], 'lower');
		expect(lower.tableData!.rows[0]!.cells[0]!.text).toBe('hello world');

		const cap = tableCellsChangeCasePatch(table, [{ row: 0, col: 0 }], 'capitalize');
		expect(cap.tableData!.rows[0]!.cells[0]!.text).toBe('Hello World');

		const toggle = tableCellsChangeCasePatch(table, [{ row: 0, col: 1 }], 'toggle');
		expect(toggle.tableData!.rows[0]!.cells[1]!.text).toBe('sECOND CELL');

		const sentence = tableCellsChangeCasePatch(next, [{ row: 0, col: 0 }], 'sentence');
		expect(sentence.tableData!.rows[0]!.cells[0]!.text).toBe('Hello world');
	});
});
