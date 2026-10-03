import { inspectXlsx, createXlsx, readXlsxRange, setXlsxCells } from 'ooxml-core/automation';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { createFileOperations } from 'ooxml-core/automation/node';

export function registerTools(server, options = {}) {
	const files = createFileOperations(options.rootDir);
	const path = z.string().min(1);
	const fileSchema = { filePath: path };
	const editSchema = { ...fileSchema, outputPath: path.optional() };
	const register = (name, description, inputSchema, readOnlyHint, handler) =>
		server.registerTool(
			name,
			{
				description,
				inputSchema,
				annotations: { readOnlyHint, destructiveHint: !readOnlyHint, openWorldHint: false },
			},
			async (params) => {
				try {
					const result = await handler(params);
					return { content: [{ type: 'text', text: JSON.stringify(result) }] };
				} catch (error) {
					return {
						isError: true,
						content: [
							{ type: 'text', text: error instanceof Error ? error.message : String(error) },
						],
					};
				}
			},
		);

	register(
		'xlsx_inspect',
		'List spreadsheet sheets, indexes, properties and warnings.',
		fileSchema,
		true,
		(p) => files.inspect(p.filePath, ['.xlsx', '.xlsm'], inspectXlsx),
	);
	register(
		'xlsx_create',
		'Create a workbook at a new path. Existing files are never overwritten.',
		{ ...fileSchema, sheets: z.array(path).max(100).default([]) },
		false,
		(p) => files.create(p.filePath, ['.xlsx'], () => createXlsx(p.sheets)),
	);
	register(
		'xlsx_read_range',
		'Read a bounded A1 range including cell values and formulas (at most 10000 cells).',
		{ ...fileSchema, sheetIndex: z.number().int().min(0), range: path },
		true,
		(p) =>
			files.inspect(p.filePath, ['.xlsx', '.xlsm'], (bytes) =>
				readXlsxRange(bytes, p.sheetIndex, p.range),
			),
	);
	register(
		'xlsx_set_cells',
		'Set cell inputs through core editing and formula recalculation. Inputs beginning with = are formulas. Without outputPath the source is updated.',
		{
			...editSchema,
			sheetIndex: z.number().int().min(0),
			cells: z
				.array(z.object({ address: path, input: z.string().max(1000000) }))
				.min(1)
				.max(10000),
		},
		false,
		(p) =>
			files.edit(
				p.filePath,
				['.xlsx', '.xlsm'],
				(bytes) => setXlsxCells(bytes, p.sheetIndex, p.cells),
				p.outputPath,
			),
	);
}

export function createServer(options = {}) {
	const server = new McpServer({ name: 'xlsx-viewer-tools', version: '0.1.0' });
	registerTools(server, options);
	return server;
}
