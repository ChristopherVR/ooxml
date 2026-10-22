import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../src/index.js';

test('xlsx MCP discovery, input validation and tool errors', async () => {
	const root = await mkdtemp(join(tmpdir(), 'xlsx-mcp-'));
	const server = createServer({ rootDir: root });
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	try {
		await server.connect(b);
		await client.connect(a);
		const { tools } = await client.listTools();
		assert.ok(tools.length >= 2);
		assert.ok(tools.every((tool) => tool.name.startsWith('xlsx_')));
		for (const name of [
			'xlsx_read_range',
			'xlsx_read_filled_cells',
			'xlsx_find',
			'xlsx_read_table',
		])
			assert.equal(tools.find((tool) => tool.name === name)?.annotations.readOnlyHint, true, name);
		const inspect = tools.find((tool) => tool.name === 'xlsx_inspect');
		assert.equal(inspect.annotations.readOnlyHint, true);
		const missing = await client.callTool({
			name: 'xlsx_inspect',
			arguments: { filePath: 'missing.xlsx' },
		});
		assert.equal(missing.isError, true);
		const invalid = await client.callTool({ name: 'xlsx_inspect', arguments: { filePath: 123 } });
		assert.equal(invalid.isError, true);
		const created = await client.callTool({
			name: 'xlsx_create',
			arguments: { filePath: 'new.xlsx' },
		});
		assert.ok(!created.isError);
		const duplicate = await client.callTool({
			name: 'xlsx_create',
			arguments: { filePath: 'new.xlsx' },
		});
		assert.equal(duplicate.isError, true);
		const inspected = await client.callTool({
			name: 'xlsx_inspect',
			arguments: { filePath: 'new.xlsx' },
		});
		assert.ok(!inspected.isError);
		const filled = await client.callTool({
			name: 'xlsx_read_filled_cells',
			arguments: { filePath: 'new.xlsx', maxCells: 10 },
		});
		assert.equal(JSON.parse(filled.content[0].text).truncated, false);
		const unbounded = await client.callTool({
			name: 'xlsx_read_filled_cells',
			arguments: { filePath: 'new.xlsx', maxCells: 1000000 },
		});
		assert.equal(unbounded.isError, true);
	} finally {
		await client.close();
		await server.close();
		await rm(root, { recursive: true, force: true });
	}
});

test('xlsx MCP proposes edits as a diff and applies only the approved proposal', async () => {
	const root = await mkdtemp(join(tmpdir(), 'xlsx-mcp-'));
	const server = createServer({ rootDir: root });
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	const call = async (name, args) => {
		const result = await client.callTool({ name, arguments: args });
		assert.ok(!result.isError, result.content[0].text);
		return JSON.parse(result.content[0].text);
	};
	try {
		await server.connect(b);
		await client.connect(a);
		const { tools } = await client.listTools();
		const propose = tools.find((tool) => tool.name === 'xlsx_propose_edits');
		assert.equal(propose.annotations.readOnlyHint, true);
		const apply = tools.find((tool) => tool.name === 'xlsx_apply_proposal');
		assert.equal(apply.annotations.readOnlyHint, false);
		await call('xlsx_create', { filePath: 'book.xlsx', sheets: ['Data'] });
		const proposed = await call('xlsx_propose_edits', {
			filePath: 'book.xlsx',
			edits: [{ op: 'setCells', sheet: 0, cells: [{ address: 'A1', input: '7' }] }],
		});
		assert.match(proposed.report, /A1 entered value: .* -> 7/);
		const unchanged = await call('xlsx_read_range', {
			filePath: 'book.xlsx',
			sheetIndex: 0,
			range: 'A1',
		});
		assert.equal(unchanged.cells[0].value, null);
		await call('xlsx_apply_proposal', { filePath: 'book.xlsx', proposal: proposed.proposal });
		const applied = await call('xlsx_read_range', {
			filePath: 'book.xlsx',
			sheetIndex: 0,
			range: 'A1',
		});
		assert.equal(applied.cells[0].value, 7);
		const stale = await client.callTool({
			name: 'xlsx_apply_proposal',
			arguments: { filePath: 'book.xlsx', proposal: proposed.proposal },
		});
		assert.equal(stale.isError, true);
		assert.match(stale.content[0].text, /document changed/);
	} finally {
		await client.close();
		await server.close();
		await rm(root, { recursive: true, force: true });
	}
});
