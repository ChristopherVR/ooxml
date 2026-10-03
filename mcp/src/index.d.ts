import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
export type OfficeFormat = 'docx' | 'xlsx' | 'visio' | 'pptx';
export interface ServerOptions {
	rootDir?: string;
	formats?: OfficeFormat[];
	registrations?: Partial<
		Record<
			OfficeFormat,
			{
				registerTools(server: McpServer, options: { rootDir?: string }): void;
			}
		>
	>;
}
export function createServer(options?: ServerOptions): Promise<McpServer>;
