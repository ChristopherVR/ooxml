import { createXlsx, setXlsxCells } from './xlsx';
import { applyXlsxProposal, proposeXlsxEdits } from './xlsx-proposal';
import type { XlsxEdit } from './xlsx-edits';
import { createEditProposal } from './proposal';
import { loadXlsx } from '../xlsx/read/load';
import { getCell } from '../xlsx/cells';

async function workbook() {
	const bytes = await createXlsx(['Data']);
	return (
		await setXlsxCells(bytes, 0, [
			{ address: 'A1', input: '1' },
			{ address: 'B1', input: '=A1*2' },
		])
	).bytes;
}

test('a proposal reports the diff of its edits and saves nothing', async () => {
	const bytes = await workbook();
	const original = bytes.slice();
	const edits: XlsxEdit[] = [
		{ op: 'setCells', sheet: 0, cells: [{ address: 'A1', input: '5' }] },
		{ op: 'renameSheet', sheet: 0, name: 'Inputs' },
		{ op: 'addSheet', name: 'Notes' },
		{ op: 'merge', sheet: 0, range: 'D1:E1' },
	];
	const { proposal, comparison, report } = await proposeXlsxEdits(bytes, edits);
	expect(bytes).toEqual(original);
	expect(proposal).toMatchObject({ format: 'xlsx', edits });
	expect(comparison.workbook).toEqual(
		expect.arrayContaining([
			{ kind: 'sheetRenamed', from: 'Data', to: 'Inputs' },
			expect.objectContaining({ kind: 'sheetAdded', sheet: 'Notes' }),
		]),
	);
	const cells = comparison.sheets[0]!.cells.map((cell) => [cell.kind, cell.after.value]);
	expect(cells).toEqual([
		['enteredValue', 5],
		['calculatedValue', 10],
	]);
	expect(report).toContain('A1 entered value: 1 -> 5');
	expect(report).toContain('Merge added D1:E1');
});

test('applying a proposal saves exactly the reviewed edits', async () => {
	const bytes = await workbook();
	const { proposal, report } = await proposeXlsxEdits(bytes, [
		{ op: 'setCells', sheet: 0, cells: [{ address: 'C3', input: 'new' }] },
		{ op: 'setName', name: 'Rate', formula: '=Data!$A$1' },
	]);
	const applied = await applyXlsxProposal(bytes, JSON.parse(JSON.stringify(proposal)));
	expect(applied.report).toBe(report);
	const saved = await loadXlsx(applied.bytes);
	expect(getCell(saved.sheets[0]!, 2, 2)?.value).toBe('new');
	expect(saved.definedNames).toEqual([{ name: 'Rate', formula: 'Data!$A$1' }]);
});

test('a proposal does not apply to a changed file or after it was altered', async () => {
	const bytes = await workbook();
	const { proposal } = await proposeXlsxEdits(bytes, [
		{ op: 'setCells', sheet: 0, cells: [{ address: 'A1', input: '2' }] },
	]);
	const changed = (await setXlsxCells(bytes, 0, [{ address: 'A2', input: 'x' }])).bytes;
	await expect(applyXlsxProposal(changed, proposal)).rejects.toThrow('document changed');
	const altered = structuredClone(proposal);
	altered.edits[0] = { op: 'deleteSheet', sheet: 0 };
	await expect(applyXlsxProposal(bytes, altered)).rejects.toThrow('changed after it was made');
	await expect(applyXlsxProposal(bytes, { ...proposal, format: 'docx' })).rejects.toThrow(
		'for docx',
	);
});

test('the seal ignores key order and undefined keys', () => {
	const bytes = new Uint8Array([1, 2, 3]);
	const a = createEditProposal('xlsx', bytes, [{ op: 'addSheet', name: 'A', at: undefined }]);
	const b = createEditProposal('xlsx', bytes, [{ name: 'A', op: 'addSheet' }]);
	expect(a.id).toBe(b.id);
	expect(a.base).toBe('039058c6f2c0cb492c533b0a4d14ef77cc0f78abccced5287d84a1a2011cfb81');
});

test('row and column edits, formats and clears', async () => {
	const bytes = await workbook();
	const { comparison } = await proposeXlsxEdits(bytes, [
		{ op: 'insertRows', sheet: 0, row: 1, count: 1 },
		{ op: 'format', sheet: 0, range: 'A2', style: { numFmt: '0%' } },
		{ op: 'insertColumns', sheet: 0, column: 'A', count: 1 },
		{ op: 'deleteColumns', sheet: 0, column: 'A', count: 1 },
		// The shifted formula is cleared again, so B2 ends as it began: empty.
		{ op: 'clear', sheet: 0, range: 'B2' },
	]);
	const kinds = comparison.sheets[0]!.cells.map((cell) => `${cell.kind} ${cell.address.row}`);
	expect(kinds).toEqual(['enteredValue 0', 'formula 0', 'enteredValue 1', 'format 1']);
});

test('invalid edits name the edit and leave the file alone', async () => {
	const bytes = await workbook();
	const cases: [unknown, string][] = [
		[{ op: 'setCells', sheet: 3, cells: [] }, 'Edit 1: Sheet does not exist'],
		[{ op: 'setCells', sheet: 0, cells: [{ address: 'ZZZZ1', input: '' }] }, 'Invalid cell'],
		[{ op: 'clear', sheet: 0, range: 'A1', what: 'everything' }, 'Cannot clear'],
		[{ op: 'format', sheet: 0, range: 'A1', style: { colour: 'red' } }, 'Unknown style keys'],
		[{ op: 'insertRows', sheet: 0, row: 0, count: 1 }, 'Invalid row'],
		[{ op: 'deleteColumns', sheet: 0, column: 'A1', count: 1 }, 'Invalid column'],
		[{ op: 'renameSheet', sheet: 0, name: 'a/b' }, 'cannot contain'],
		[{ op: 'deleteSheet', sheet: 0 }, 'at least one sheet'],
		[{ op: 'explode' }, 'Unknown edit: explode'],
	];
	for (const [edit, message] of cases)
		await expect(proposeXlsxEdits(bytes, [edit as XlsxEdit])).rejects.toThrow(message);
	await expect(
		proposeXlsxEdits(bytes, [
			{ op: 'setCells', sheet: 0, cells: [{ address: 'A1', input: '9' }] },
			{ op: 'addSheet', name: 'data' },
		]),
	).rejects.toThrow('Edit 2: That name is already taken.');
});
