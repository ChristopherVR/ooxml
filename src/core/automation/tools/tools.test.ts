import { automationTools, findAutomationTool, UNTRUSTED_NOTE, type AutomationTool } from './index';
import { createXlsxFromSheets } from '../outline';

const all = Object.values(automationTools).flat();

function tool(name: string): AutomationTool {
	const found = findAutomationTool(name);
	if (!found) throw new Error(`No tool ${name}`);
	return found;
}

async function read(name: string, bytes: Uint8Array, params: unknown) {
	const found = tool(name);
	if (found.kind !== 'read') throw new Error(`${name} is not a read tool`);
	return found.run(bytes, params) as Promise<Record<string, unknown>>;
}

test('tool names are unique, prefixed by their format and described honestly', () => {
	expect(new Set(all.map((t) => t.name)).size).toBe(all.length);
	for (const t of all) {
		expect(t.name.startsWith(`${t.format}_`)).toBe(true);
		expect(t.extensions.length).toBeGreaterThan(0);
		expect(t.description.endsWith(UNTRUSTED_NOTE)).toBe(t.kind === 'read');
		expect(t.jsonSchema()).toMatchObject({ type: 'object' });
	}
	expect(findAutomationTool('missing')).toBeUndefined();
});

test('the JSON Schema states the parameters a model must send', () => {
	const schema = tool('xlsx_read_range').jsonSchema();
	expect(schema).toMatchObject({
		properties: { sheetIndex: { type: 'integer' }, range: { type: 'string' } },
		required: ['sheetIndex', 'range'],
		additionalProperties: false,
	});
});

test('parameters are validated in core, not only by the transport', async () => {
	const bytes = await createXlsxFromSheets([{ name: 'Data', rows: [['a']] }]);
	await expect(read('xlsx_read_range', bytes, { sheetIndex: -1, range: 'A1' })).rejects.toThrow(
		/Invalid xlsx_read_range parameters/,
	);
	await expect(read('xlsx_inspect', bytes, { unexpected: true })).rejects.toThrow(/unexpected/i);
});

test('spreadsheet reads stay bounded on a large sheet', async () => {
	const rows = Array.from({ length: 300 }, (_, row) =>
		Array.from({ length: 10 }, (_, col) => row * 10 + col),
	);
	const bytes = await createXlsxFromSheets([
		{ name: 'Big', rows },
		{ name: 'Small', rows: [['total', '=SUM(Big!A1:A3)']] },
	]);
	const inspected = await read('xlsx_inspect', bytes, {});
	expect(inspected.sheets).toHaveLength(2);

	const filled = await read('xlsx_read_filled_cells', bytes, {});
	expect(filled.truncated).toBe(true);
	const [big] = filled.sheets as { cells: unknown[]; total: number; omitted: number }[];
	expect(big!.cells).toHaveLength(2000);
	expect(big!.total).toBe(3000);
	expect(big!.omitted).toBe(1000);

	const small = await read('xlsx_read_filled_cells', bytes, { sheetIndex: 1, maxCells: 10 });
	expect(small).toMatchObject({
		truncated: false,
		sheets: [{ index: 1, name: 'Small', total: 2, omitted: 0 }],
	});
	await expect(read('xlsx_read_filled_cells', bytes, { maxCells: 10_001 })).rejects.toThrow();
	await expect(read('xlsx_read_filled_cells', bytes, { sheetIndex: 5 })).rejects.toThrow(
		'Invalid sheet',
	);

	const found = await read('xlsx_find', bytes, {
		text: '29?',
		sheetIndex: 0,
		wholeCell: true,
		maxResults: 3,
	});
	expect(found).toMatchObject({ total: 10, truncated: true });
	expect(found.matches).toHaveLength(3);
	await expect(
		read('xlsx_read_range', bytes, { sheetIndex: 0, range: 'A1:Z1000' }),
	).rejects.toThrow('at most 10000');
});

test('Word tools inspect, edit and suggest through the same definitions', async () => {
	const create = tool('docx_create');
	if (create.kind !== 'create') throw new Error('docx_create is not a create tool');
	const bytes = await create.run({ paragraphs: ['Hello'] });
	const inspected = await read('docx_inspect', bytes, {});
	const paragraphId = (inspected.blocks as { id: string }[])[0]!.id;
	const set = tool('docx_set_run_text');
	if (set.kind !== 'edit') throw new Error('docx_set_run_text is not an edit tool');
	const edited = await set.run(bytes, { paragraphId, runIndex: 0, text: 'Bye' });
	const after = await read('docx_inspect', edited.bytes, {});
	expect(after.blocks).toMatchObject([{ runs: [{ text: 'Bye' }] }]);
	const suggest = tool('docx_suggest_run_text');
	if (suggest.kind !== 'edit') throw new Error('docx_suggest_run_text is not an edit tool');
	const suggested = await suggest.run(bytes, {
		suggestions: [{ paragraphId, runIndex: 0, text: 'Hi' }],
		author: 'Reviewer',
	});
	const redlined = await read('docx_inspect', suggested.bytes, {});
	expect(redlined.blocks).toMatchObject([
		{
			runs: [
				{ text: 'Hello', revision: { kind: 'delete', author: 'Reviewer' } },
				{ text: 'Hi', revision: { kind: 'insert', author: 'Reviewer' } },
			],
		},
	]);
	expect(await read('docx_list_form_fields', bytes, {})).toMatchObject({ fields: [] });
});
