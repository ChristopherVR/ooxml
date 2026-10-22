import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerAutomationTools } from 'ooxml-core/automation/node';
import { pptxTools } from 'ooxml-core/pptx/automation';

/** Optional configuration for the MCP server. */
export interface McpContextOptions {
	/**
	 * Root directory under which all `filePath` arguments are scoped.
	 * Defaults to `process.env.PPTX_TOOLS_ROOT` if set, else `process.cwd()` with a warning.
	 */
	rootDir?: string;
}

export function createServer(options?: McpContextOptions): McpServer {
	const server = new McpServer({ name: 'pptx-viewer-tools', version: '1.0.0' });
	registerTools(server, options);
	return server;
}

/**
 * Registers the core's PowerPoint tools as file tools under the root. The tools (names,
 * schemas, operations) are `pptxTools` in `ooxml-core/pptx/automation`, registered the same way
 * as the Word, Excel and Visio tools.
 */
export function registerTools(server: McpServer, options?: McpContextOptions): void {
	let rootDir = options?.rootDir ?? process.env['PPTX_TOOLS_ROOT'];
	if (!rootDir) {
		process.stderr.write(
			'pptx-viewer-tools: no root configured (--root or PPTX_TOOLS_ROOT); using the current working directory\n',
		);
		rootDir = process.cwd();
	}
	registerAutomationTools(server, pptxTools, { rootDir });
}
