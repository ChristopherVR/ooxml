import JSZip from 'jszip';
import { createDocx, inspectDocx, setDocxRunText } from './docx.js';
import { createXlsx, readXlsxRange, setXlsxCells } from './xlsx.js';
import { createDocument } from '../docx/model.js';
import { saveDocx } from '../docx/save.js';
import { loadDocx } from '../docx/parse.js';
import { fixture, shape } from '../visio/test-fixtures.js';
import { editVisio, inspectVisio } from './index.js';

test('Word edits retain formatting and unmodeled package parts on reload', async () => {
	const model = createDocument();
	const paragraph = model.blocks[0]!;
	if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
	paragraph.runs = [
		{ text: 'before', bold: true },
		{ text: ' untouched', italic: true },
	];
	const zip = await JSZip.loadAsync(await saveDocx(model));
	zip.file('customXml/item1.xml', '<custom>preserve me</custom>');
	const input = await zip.generateAsync({ type: 'uint8array' });
	const loaded = await inspectDocx(input);
	const edited = await setDocxRunText(input, loaded.blocks[0]!.id, 0, 'after');
	const reloaded = await loadDocx(edited.bytes);
	expect(reloaded.model.blocks[0]).toMatchObject({
		runs: [
			{ text: 'after', bold: true },
			{ text: ' untouched', italic: true },
		],
	});
	expect(
		await (await JSZip.loadAsync(edited.bytes)).file('customXml/item1.xml')!.async('string'),
	).toBe('<custom>preserve me</custom>');
});

test('Word rejects invalid targets, fields and tracked edits', async () => {
	const bytes = await createDocx(['Hello']);
	await expect(setDocxRunText(bytes, 'missing', 0, 'change')).rejects.toThrow('does not exist');
	const model = createDocument();
	const paragraph = model.blocks[0]!;
	if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
	paragraph.runs = [{ text: '2026', field: { instr: 'DATE', simple: true } }];
	const fieldBytes = await saveDocx(model);
	const id = (await inspectDocx(fieldBytes)).blocks[0]!.id;
	await expect(setDocxRunText(fieldBytes, id, 0, 'change')).rejects.toThrow('ordinary');
	model.trackChanges = true;
	paragraph.runs = [{ text: 'tracked' }];
	const tracked = await saveDocx(model);
	await expect(
		setDocxRunText(tracked, (await inspectDocx(tracked)).blocks[0]!.id, 0, 'change'),
	).rejects.toThrow('ordinary');
});

test('spreadsheet batch edits use the core formula engine and save cached results', async () => {
	const input = await createXlsx(['Data', 'Summary']);
	const edited = await setXlsxCells(input, 0, [
		{ address: 'A1', input: '3' },
		{ address: 'A2', input: '4' },
		{ address: 'A3', input: '=SUM(A1:A2)' },
	]);
	const result = await readXlsxRange(edited.bytes, 0, 'A1:A3');
	expect(result.cells.map((cell) => cell.value)).toEqual([3, 4, 7]);
	expect(result.cells[2]!.formula).toBe('SUM(A1:A2)');
	await expect(readXlsxRange(input, 0, 'A:XFD')).rejects.toThrow('10000');
	await expect(setXlsxCells(input, 0, [{ address: 'XFE1', input: 'bad' }])).rejects.toThrow(
		'Invalid',
	);
	await expect(createXlsx(['same', 'SAME'])).rejects.toThrow();
});

test('Visio automation delegates source-backed edits and returns core diagnostics', async () => {
	const input = await fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', '<Text>before</Text>')}</Shapes>` }],
	});
	const result = await editVisio(input, [
		{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'after' },
	]);
	expect((await inspectVisio(result.bytes)).pages[0]!.shapes[0]).toMatchObject({ text: 'after' });
	expect(result.diagnostics).toContainEqual(
		expect.objectContaining({ code: 'edit-caches-not-recalculated' }),
	);
});
