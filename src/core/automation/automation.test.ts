import JSZip from 'jszip';
import { createDocx, inspectDocx, setDocxRunText, suggestDocxRunText } from './docx';
import { acceptAllRevisions, listRevisions, rejectAllRevisions } from '../docx/revision-commands';
import { createXlsx, readXlsxFilledCells, readXlsxRange, setXlsxCells } from './xlsx';
import { createDocument } from '../docx/model';
import { saveDocx } from '../docx/save';
import { loadDocx } from '../docx/parse';
import { fixture, shape } from '../visio/test-fixtures';
import { editVisio, inspectVisio } from './index';

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

test('filled-cell reads reach past any fixed range and report what the cap left out', async () => {
	const input = await createXlsx(['Data', 'Far']);
	let bytes = (
		await setXlsxCells(input, 0, [
			{ address: 'A1', input: 'Item' },
			{ address: 'AZ500', input: '=1+1' },
		])
	).bytes;
	bytes = (await setXlsxCells(bytes, 1, [{ address: 'C3', input: 'x' }])).bytes;
	const all = await readXlsxFilledCells(bytes);
	expect(all.truncated).toBe(false);
	expect(all.sheets[0]!.cells).toEqual([
		{ address: 'A1', value: 'Item' },
		{ address: 'AZ500', value: 2, formula: '1+1' },
	]);
	expect(all.sheets[1]!.cells.map((c) => c.address)).toEqual(['C3']);
	const capped = await readXlsxFilledCells(bytes, { maxCells: 1 });
	expect(capped.truncated).toBe(true);
	expect(capped.sheets.map((s) => [s.cells.length, s.omitted])).toEqual([
		[1, 1],
		[0, 1],
	]);
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

test('Word suggestions are tracked changes that accept to the new text and reject to the old', async () => {
	const input = await createDocx(['Launch brief', 'Ships in November.']);
	const id = (await inspectDocx(input)).blocks[1]!;
	if (id.type !== 'paragraph') throw new Error('expected a paragraph');
	const { bytes } = await suggestDocxRunText(input, [
		{ paragraphId: id.id, runIndex: 0, text: 'Ships in December.' },
	]);
	const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:del [^>]*w:author="Assistant"/u);
	expect(xml).toMatch(/<w:ins [^>]*w:author="Assistant"/u);
	const { model } = await loadDocx(bytes);
	expect(
		listRevisions(model)
			.map((r) => r.kind)
			.sort(),
	).toEqual(['delete', 'insert']);
	const text = (m: typeof model) => {
		const block = m.blocks[1]!;
		return block.type === 'paragraph' ? block.runs.map((r) => r.text).join('') : '';
	};
	expect(text(acceptAllRevisions(model))).toBe('Ships in December.');
	expect(text(rejectAllRevisions(model))).toBe('Ships in November.');
	await expect(
		suggestDocxRunText(input, [{ paragraphId: 'missing', runIndex: 0, text: 'x' }]),
	).rejects.toThrow('does not exist');
});
