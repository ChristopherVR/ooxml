import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

const modules = {
	docx: 'docx-viewer-mcp',
	xlsx: 'xlsx-viewer-mcp',
	visio: 'visio-viewer-mcp',
	pptx: 'pptx-viewer-mcp/mcp',
};

/** Compose tools owned by the viewer repositories; contains no document logic. */
export async function createServer(options = {}) {
	const formats = options.formats ?? Object.keys(modules);
	if (!formats.length || new Set(formats).size !== formats.length)
		throw new Error('Select at least one format, with no duplicates');
	const registrations = await Promise.all(
		formats.map(async (format) => {
			if (!Object.hasOwn(modules, format)) throw new Error(`Unknown format: ${format}`);
			const implementation =
				options.registrations?.[format] ??
				(await import(modules[format]).catch((error) => {
					throw new Error(
						`Install a compatible ${format}-viewer-mcp package before enabling ${format}`,
						{ cause: error },
					);
				}));
			if (typeof implementation.registerTools !== 'function')
				throw new Error(`${format}-viewer-mcp must export registerTools; update that package`);
			return implementation.registerTools;
		}),
	);
	const server = new McpServer({ name: 'ooxml-tools', version: '0.1.0' });
	for (const register of registrations) register(server, { rootDir: options.rootDir });
	return server;
}
