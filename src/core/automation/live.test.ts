import { createLiveToolHost, liveEditorFromElement, type LiveEditor } from './live';
import { createXlsxFromSheets } from './outline';
import { automationTools } from './tools/index';

const tools = Object.values(automationTools).flat();

/** An editor held in memory, with a hook to simulate the user typing mid-call. */
function fakeEditor(format: LiveEditor['format'], initial: Uint8Array) {
	let bytes = initial;
	const editor = {
		format,
		readOnly: false,
		loads: 0,
		save: () => bytes,
		load(next: Uint8Array) {
			bytes = next;
			editor.loads += 1;
		},
		type(next: Uint8Array) {
			bytes = next;
		},
	};
	return editor;
}

async function book(value: string) {
	return createXlsxFromSheets([{ name: 'Data', rows: [[value]] }]);
}

const readA1 = { sheetIndex: 0, range: 'A1' };

test('a call must name a registered document; there is no default', async () => {
	const host = createLiveToolHost(tools);
	host.register('one', fakeEditor('xlsx', await book('a')));
	await expect(host.call('xlsx_inspect', {} as never)).rejects.toThrow('needs a target');
	await expect(host.call('xlsx_inspect', { target: 'two' })).rejects.toThrow('"two"');
	await expect(host.call('no_such_tool', { target: 'one' })).rejects.toThrow('Unknown tool');
});

test('reads see the unsaved state of the open editor', async () => {
	const host = createLiveToolHost(tools);
	const editor = fakeEditor('xlsx', await book('before'));
	host.register('sheet', editor);
	editor.type(await book('typed since load'));
	const { result, changed } = await host.call('xlsx_read_range', {
		target: 'sheet',
		params: readA1,
	});
	expect(changed).toBe(false);
	expect(result).toMatchObject({ cells: [{ value: 'typed since load' }] });
});

test('an edit is loaded back into the editor it targeted, and only that one', async () => {
	const host = createLiveToolHost(tools);
	const first = fakeEditor('xlsx', await book('one'));
	const second = fakeEditor('xlsx', await book('two'));
	host.register('first', first);
	host.register('second', second);
	const outcome = await host.call('xlsx_set_cells', {
		target: 'second',
		params: { sheetIndex: 0, cells: [{ address: 'A1', input: 'edited' }] },
	});
	expect(outcome.changed).toBe(true);
	expect([first.loads, second.loads]).toEqual([0, 1]);
	const reread = await host.call('xlsx_read_range', { target: 'second', params: readA1 });
	expect(reread.result).toMatchObject({ cells: [{ value: 'edited' }] });
	const untouched = await host.call('xlsx_read_range', { target: 'first', params: readA1 });
	expect(untouched.result).toMatchObject({ cells: [{ value: 'one' }] });
});

test('a tool for another format is refused before it touches the document', async () => {
	const host = createLiveToolHost(tools);
	const editor = fakeEditor('xlsx', await book('a'));
	host.register('sheet', editor);
	await expect(host.call('docx_inspect', { target: 'sheet' })).rejects.toThrow(
		'works on docx documents',
	);
	expect(editor.loads).toBe(0);
});

test('a read-only editor is never edited, but can be read', async () => {
	const host = createLiveToolHost(tools);
	const editor = fakeEditor('xlsx', await book('a'));
	editor.readOnly = true;
	host.register('sheet', editor);
	const edit = { sheetIndex: 0, cells: [{ address: 'A1', input: 'x' }] };
	await expect(host.call('xlsx_set_cells', { target: 'sheet', params: edit })).rejects.toThrow(
		'read-only',
	);
	await expect(
		host.call('xlsx_read_range', { target: 'sheet', params: readA1 }),
	).resolves.toBeDefined();
	expect(host.targets()).toEqual([{ id: 'sheet', format: 'xlsx', readOnly: true }]);
});

test('review can discard an edit, and sees the file before and after', async () => {
	const seen: { before: number; after: number }[] = [];
	let accept = false;
	const host = createLiveToolHost(tools, {
		review: ({ before, after }) => {
			seen.push({ before: before.length, after: after.length });
			return accept;
		},
	});
	const editor = fakeEditor('xlsx', await book('a'));
	host.register('sheet', editor);
	const params = { sheetIndex: 0, cells: [{ address: 'A1', input: 'b' }] };
	const refused = await host.call('xlsx_set_cells', { target: 'sheet', params });
	expect(refused).toMatchObject({ changed: false, result: { discarded: true } });
	expect(editor.loads).toBe(0);
	accept = true;
	expect((await host.call('xlsx_set_cells', { target: 'sheet', params })).changed).toBe(true);
	expect(seen).toHaveLength(2);
	expect(seen[0]!.before).toBeGreaterThan(0);
});

test('an edit is not applied over what the user typed while it was prepared', async () => {
	const editor = fakeEditor('xlsx', await book('a'));
	const typed = await book('typed meanwhile');
	const host = createLiveToolHost(tools, {
		review: () => {
			editor.type(typed);
			return true;
		},
	});
	host.register('sheet', editor);
	await expect(
		host.call('xlsx_set_cells', {
			target: 'sheet',
			params: { sheetIndex: 0, cells: [{ address: 'A1', input: 'b' }] },
		}),
	).rejects.toThrow('changed while the edit was prepared');
	expect(editor.loads).toBe(0);
});

test('describe lists only tools for open formats and requires a target', async () => {
	const host = createLiveToolHost(tools);
	expect(host.describe()).toEqual([]);
	host.register('sheet', fakeEditor('xlsx', await book('a')));
	const described = host.describe();
	expect(described.length).toBeGreaterThan(0);
	expect(described.every((tool) => tool.format === 'xlsx')).toBe(true);
	// Create tools have no document to target, so they are not offered.
	expect(described.some((tool) => tool.name === 'xlsx_create')).toBe(false);
	for (const tool of described)
		expect(tool.inputSchema).toMatchObject({ required: ['target'], type: 'object' });
});

test('ids are unique and unregistering frees them', async () => {
	const host = createLiveToolHost(tools);
	const stop = host.register('a', fakeEditor('xlsx', await book('a')));
	expect(() => host.register('a', fakeEditor('xlsx', new Uint8Array()))).toThrow('already');
	stop();
	expect(host.targets()).toEqual([]);
	expect(() => host.register('a', fakeEditor('xlsx', new Uint8Array()))).not.toThrow();
});

test('overlapping edits to one document run one after the other', async () => {
	const host = createLiveToolHost(tools);
	const editor = fakeEditor('xlsx', await book('a'));
	host.register('sheet', editor);
	const set = (address: string, input: string) =>
		host.call('xlsx_set_cells', {
			target: 'sheet',
			params: { sheetIndex: 0, cells: [{ address, input }] },
		});
	await Promise.all([set('A1', 'x'), set('B1', 'y')]);
	const { result } = await host.call('xlsx_read_range', {
		target: 'sheet',
		params: { sheetIndex: 0, range: 'A1:B1' },
	});
	expect(result).toMatchObject({ cells: [{ value: 'x' }, { value: 'y' }] });
});

test('elements adapt by format: Visio exports, the rest save bytes', async () => {
	const bytes = new Uint8Array([1, 2, 3]);
	const loads: unknown[] = [];
	const sheet = liveEditorFromElement(
		'xlsx',
		{ saveBytes: () => bytes, load: (b, name) => loads.push([b, name]), readOnly: true },
		'Book.xlsx',
	);
	expect(await sheet.save()).toBe(bytes);
	await sheet.load(bytes);
	expect(loads).toEqual([[bytes, 'Book.xlsx']]);
	expect(sheet.readOnly).toBe(true);

	const visio = liveEditorFromElement('visio', {
		exportVsdx: async () => ({ bytes }),
		load: () => undefined,
	});
	expect(await visio.save()).toBe(bytes);
	expect(visio.readOnly).toBe(false);

	const editing = liveEditorFromElement('xlsx', {
		saveBytes: () => bytes,
		load: () => undefined,
		commitEdit: () => false,
	});
	await expect(Promise.resolve().then(() => editing.save())).rejects.toThrow('cell edit');
});
