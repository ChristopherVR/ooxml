import { eighthPoints, signedTwips, twips } from './units';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx, type DocumentModel, type Table } from './index';
import { at, must } from './test-support/access';

const para = (id: string) => ({ type: 'paragraph' as const, id, runs: [{ text: id }] });

function modelWith(table: Omit<Table, 'type' | 'id'>): DocumentModel {
	const model = createDocument();
	model.blocks = [para('p1'), { type: 'table', id: 't1', ...table }, para('p2')];
	return model;
}
async function tableXml(model: DocumentModel): Promise<string> {
	const xml = await (
		await JSZip.loadAsync(await saveDocx(model))
	)
		.file('word/document.xml')!
		.async('string');
	return /<w:tbl>[\s\S]*<\/w:tbl>/.exec(xml)![0];
}

describe('new table cell properties', () => {
	it('writes tcPr children in CT_TcPr order', async () => {
		const xml = await tableXml(
			modelWith({
				rows: [
					[
						{
							paragraphs: [para('a')],
							widthTwips: twips(3000),
							gridSpan: 2,
							verticalMerge: 'restart',
							borders: {
								top: { style: 'double', sizeEighthPoints: eighthPoints(6), color: '#112233' },
							},
							shadingFill: '#D9D9D9',
							margins: { top: twips(10), left: twips(20), bottom: twips(30), right: twips(40) },
							verticalAlign: 'center',
						},
					],
				],
			}),
		);
		const tcPr = must(/<w:tcPr>([\s\S]*?)<\/w:tcPr>/.exec(xml)?.[1], 'tcPr contents');
		const order = [
			...tcPr.matchAll(/<w:(tcW|gridSpan|vMerge|tcBorders|shd|tcMar|vAlign)[ >/]/g),
		].map((match) => match[1]);
		expect(order).toEqual(['tcW', 'gridSpan', 'vMerge', 'tcBorders', 'shd', 'tcMar', 'vAlign']);
		expect(tcPr).toContain('<w:gridSpan w:val="2"/>');
		expect(tcPr).toContain('<w:vMerge w:val="restart"/>');
		expect(tcPr).toContain('w:fill="D9D9D9"');
		expect(tcPr).toContain('<w:vAlign w:val="center"/>');
		expect(tcPr).toMatch(/<w:tcW w:w="3000" w:type="dxa"\/>/);
	});

	it('gives cells without a width their grid width, so Word does not shrink them to content', async () => {
		const xml = await tableXml(
			modelWith({
				grid: [twips(1000), twips(2000), twips(3000)],
				rows: [
					[{ paragraphs: [para('a')], gridSpan: 2 }, { paragraphs: [para('b')] }],
					[
						{ paragraphs: [para('c')], widthTwips: twips(900) },
						{ paragraphs: [para('d')] },
						{ paragraphs: [para('e')], widthPercent: 0.2 },
					],
				],
			}),
		);
		const widths = [...xml.matchAll(/<w:tcW w:w="(\d+)" w:type="dxa"\/>/g)].map((m) => m[1]);
		expect(widths).toEqual(['3000', '3000', '900', '2000']);
		expect(xml).toContain('<w:tcW w:w="1000" w:type="pct"/>');
	});

	it('writes percentage table widths and reads them back', async () => {
		const model = modelWith({ rows: [[{ paragraphs: [para('a')] }]], widthPercent: 0.5 });
		expect(await tableXml(model)).toContain('<w:tblW w:w="2500" w:type="pct"/>');
		const table = (await loadDocx(await saveDocx(model))).model.blocks[1] as Table;
		expect(table).toMatchObject({ widthPercent: 0.5, layout: 'autofit' });
		expect(table.widthTwips).toBeUndefined();
	});

	it('sizes tblGrid from gridSpan and marks short rows with gridAfter', async () => {
		const xml = await tableXml(
			modelWith({
				rows: [
					[{ paragraphs: [para('a')], gridSpan: 2 }, { paragraphs: [para('b')] }],
					[{ paragraphs: [para('c')] }, { paragraphs: [para('d')] }],
				],
			}),
		);
		expect(xml.match(/<w:gridCol /g)).toHaveLength(3);
		const secondRow = xml.split('<w:tr>')[2];
		expect(secondRow).toContain('<w:gridAfter w:val="1"/>');
	});

	it('round-trips written cell properties through the parser', async () => {
		const model = modelWith({
			rows: [
				[
					{
						paragraphs: [para('a')],
						gridSpan: 2,
						shadingFill: 'D9D9D9',
						verticalAlign: 'bottom',
						widthTwips: twips(4000),
					},
				],
				[{ paragraphs: [para('b')] }, { paragraphs: [para('c')] }],
			],
		});
		const table = (await loadDocx(await saveDocx(model))).model.blocks[1] as Table;
		expect(at(at(table.rows, 0), 0)).toMatchObject({
			gridSpan: 2,
			verticalAlign: 'bottom',
			widthTwips: 4000,
		});
		expect(table.grid).toHaveLength(2);
	});

	it('writes row properties, indent, default margins and exact justification', async () => {
		const xml = await tableXml(
			modelWith({
				rows: [[{ paragraphs: [para('a')] }]],
				rowProperties: [
					{ heightTwips: twips(400), heightRule: 'exact', cantSplit: true, header: true },
				],
				indentTwips: signedTwips(120),
				cellMargins: { left: twips(108), right: twips(108) },
				alignment: 'right',
				justification: 'end',
			}),
		);
		expect(xml).toContain('<w:jc w:val="end"/>');
		expect(xml).toContain('<w:tblInd w:w="120" w:type="dxa"/>');
		expect(xml).toContain('<w:tblCellMar>');
		expect(xml).toMatch(
			/<w:trPr><w:cantSplit\/><w:trHeight w:val="400" w:hRule="exact"\/><w:tblHeader\/>/,
		);
		const stale = await tableXml(
			modelWith({
				rows: [[{ paragraphs: [para('a')] }]],
				alignment: 'center',
				justification: 'end',
			}),
		);
		expect(stale).toContain('<w:jc w:val="center"/>');
	});
});
