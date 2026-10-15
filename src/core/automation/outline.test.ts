import JSZip from 'jszip';
import { loadDocx } from '../docx/parse';
import { createDocxFromOutline, createXlsxFromSheets } from './outline';
import { inspectXlsx, readXlsxRange } from './xlsx';

test('a Word outline becomes styled headings, paragraphs and real lists', async () => {
	const bytes = await createDocxFromOutline([
		{ type: 'title', text: 'Launch plan' },
		{ type: 'heading', text: 'Goals', level: 2 },
		{ type: 'paragraph', text: 'Ship in December.' },
		{ type: 'bullets', items: ['Budget', 'Team'] },
		{ type: 'numbered', items: ['Draft', 'Review'] },
	]);
	const { model } = await loadDocx(bytes);
	const paragraphs = model.blocks.flatMap((b) => (b.type === 'paragraph' ? [b] : []));
	expect(paragraphs.map((p) => [p.style, p.runs[0]?.text])).toEqual([
		['Title', 'Launch plan'],
		['Heading2', 'Goals'],
		[undefined, 'Ship in December.'],
		[undefined, 'Budget'],
		[undefined, 'Team'],
		[undefined, 'Draft'],
		[undefined, 'Review'],
	]);
	const [bullet, , numbered] = paragraphs.slice(3).map((p) => p.numbering?.numId);
	expect(bullet).toBeDefined();
	expect(numbered).not.toBe(bullet);
	expect(await (await JSZip.loadAsync(bytes)).file('word/numbering.xml')?.async('string')).toMatch(
		/w:numFmt w:val="bullet"/,
	);
});

test('a Word outline rejects unknown blocks and empty lists', async () => {
	await expect(createDocxFromOutline([{ type: 'bullets', items: [] }])).rejects.toThrow('list');
	await expect(createDocxFromOutline([{ type: 'video', text: 'x' } as never])).rejects.toThrow(
		'Unknown',
	);
});

test('a workbook outline fills each sheet from A1 and keeps formulas', async () => {
	const bytes = await createXlsxFromSheets([
		{
			name: 'Budget',
			rows: [
				['Item', 'Cost'],
				['Venue', 4200],
				['Total', '=SUM(B2:B2)'],
			],
		},
		{ name: 'Notes', rows: [[null, true]] },
	]);
	expect((await inspectXlsx(bytes)).sheets.map((s) => s.name)).toEqual(['Budget', 'Notes']);
	const cells = (await readXlsxRange(bytes, 0, 'A1:B3')).cells;
	expect(cells.find((c) => c.address === 'B2')?.value).toBe(4200);
	expect(cells.find((c) => c.address === 'B3')).toMatchObject({
		formula: 'SUM(B2:B2)',
		value: 4200,
	});
	expect((await readXlsxRange(bytes, 1, 'B1')).cells[0]?.value).toBe(true);
	await expect(createXlsxFromSheets([{ name: 'a/b', rows: [] }])).rejects.toThrow();
});
