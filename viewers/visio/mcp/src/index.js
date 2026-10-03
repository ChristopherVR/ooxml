import { inspectVisio, editVisio } from 'ooxml-core/automation';
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

	const target = { pageId: path, shapeId: path };
	const finite = z.number().finite();
	const commands = z.discriminatedUnion('type', [
		z.object({ ...target, type: z.literal('replace-plain-text'), text: z.string().max(1000000) }),
		z.object({ ...target, type: z.literal('move-shape'), x: finite, y: finite }),
		z.object({
			...target,
			type: z.literal('resize-shape'),
			width: finite.positive(),
			height: finite.positive(),
		}),
		z.object({ ...target, type: z.literal('delete-shape') }),
		z.object({
			...target,
			type: z.literal('create-rectangle'),
			x: finite,
			y: finite,
			width: finite.positive(),
			height: finite.positive(),
			text: z.string().max(1000000).optional(),
		}),
	]);
	register(
		'visio_inspect',
		'Inspect VSDX pages, shape ids, text, geometry and core diagnostics.',
		fileSchema,
		true,
		(p) => files.inspect(p.filePath, ['.vsdx'], inspectVisio),
	);
	register(
		'visio_edit',
		'Apply an atomic batch of core-supported experimental Visio text and geometry edits. Coordinates use inches with a bottom-left origin. Unsupported edits reject; diagnostics report fidelity limits. Without outputPath the source is updated.',
		{ ...editSchema, edits: z.array(commands).min(1).max(1000) },
		false,
		(p) => files.edit(p.filePath, ['.vsdx'], (bytes) => editVisio(bytes, p.edits), p.outputPath),
	);
}

export function createServer(options = {}) {
	const server = new McpServer({ name: 'visio-viewer-tools', version: '0.1.0' });
	registerTools(server, options);
	return server;
}
