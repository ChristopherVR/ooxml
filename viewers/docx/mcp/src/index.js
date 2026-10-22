import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { docxTools } from 'ooxml-core/automation';
import { registerAutomationTools } from 'ooxml-core/automation/node';

/**
 * Registers the core's docx tools as file tools under `rootDir`. The schemas and operations live
 * in `ooxml-core/automation`, shared with the suite's assistant and any host backend.
 */
export function registerTools(server, options = {}) {
	registerAutomationTools(server, docxTools, options);
}

export function createServer(options = {}) {
	const server = new McpServer({ name: 'docx-viewer-tools', version: '0.1.0' });
	registerTools(server, options);
	return server;
}
