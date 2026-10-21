import { mkdtemp, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createFileOperations, registerAutomationTools, runAutomationFileTool } from './node';
import { findAutomationTool, xlsxTools } from './tools/index';

test('file tools read, create and edit confined files', async () => {
	const root = await mkdtemp(join(tmpdir(), 'ooxml-file-tools-'));
	const tool = (name: string) => findAutomationTool(name)!;
	try {
		const files = createFileOperations(root);
		expect(
			await runAutomationFileTool(files, tool('xlsx_create'), {
				filePath: 'book.xlsx',
				sheets: ['Data'],
			}),
		).toEqual({ filePath: 'book.xlsx', saved: true });
		await expect(
			runAutomationFileTool(files, tool('xlsx_create'), { filePath: 'book.xlsx' }),
		).rejects.toThrow();
		const saved = await runAutomationFileTool(files, tool('xlsx_set_cells'), {
			filePath: 'book.xlsx',
			outputPath: 'copy.xlsx',
			sheetIndex: 0,
			cells: [{ address: 'B2', input: '=1+1' }],
		});
		expect(saved).toMatchObject({ filePath: 'copy.xlsx', saved: true });
		expect(saved).not.toHaveProperty('bytes');
		expect(
			await runAutomationFileTool(files, tool('xlsx_read_range'), {
				filePath: 'copy.xlsx',
				sheetIndex: 0,
				range: 'B2',
			}),
		).toMatchObject({ cells: [{ address: 'B2', value: 2, formula: '1+1' }] });
		await expect(
			runAutomationFileTool(files, tool('xlsx_inspect'), { filePath: '../outside.xlsx' }),
		).rejects.toThrow('outside the configured root');
		expect((await readdir(root)).sort()).toEqual(['book.xlsx', 'copy.xlsx']);
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});

test('MCP registration exposes the core tools with file paths and hints', async () => {
	const root = await mkdtemp(join(tmpdir(), 'ooxml-mcp-tools-'));
	const server = new McpServer({ name: 'test', version: '1.0.0' });
	registerAutomationTools(server, xlsxTools, { rootDir: root });
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	try {
		await server.connect(b);
		await client.connect(a);
		const { tools } = await client.listTools();
		expect(tools.map((t) => t.name)).toEqual(xlsxTools.map((t) => t.name));
		const range = tools.find((t) => t.name === 'xlsx_read_range')!;
		expect(range.annotations).toMatchObject({ readOnlyHint: true, destructiveHint: false });
		expect(range.inputSchema.required).toEqual(['filePath', 'sheetIndex', 'range']);
		const edit = tools.find((t) => t.name === 'xlsx_set_cells')!;
		expect(Object.keys(edit.inputSchema.properties ?? {})).toContain('outputPath');
		expect(edit.annotations).toMatchObject({ readOnlyHint: false, destructiveHint: true });

		const created = await client.callTool({
			name: 'xlsx_create',
			arguments: { filePath: 'new.xlsx' },
		});
		expect(created.isError).toBeFalsy();
		const filled = await client.callTool({
			name: 'xlsx_read_filled_cells',
			arguments: { filePath: 'new.xlsx', maxCells: 5 },
		});
		const [content] = filled.content as { type: string; text: string }[];
		expect(JSON.parse(content!.text)).toMatchObject({ truncated: false });
		const missing = await client.callTool({
			name: 'xlsx_inspect',
			arguments: { filePath: 'missing.xlsx' },
		});
		expect(missing.isError).toBe(true);
	} finally {
		await client.close();
		await server.close();
		await rm(root, { recursive: true, force: true });
	}
});
