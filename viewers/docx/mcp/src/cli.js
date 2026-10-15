#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './index.js';

// `--root <dir>` (or `--root=<dir>`); a bare first argument still works.
const args = process.argv.slice(2);
const flag = args.findIndex((a) => a === '--root' || a.startsWith('--root='));
const rootArg =
	flag < 0
		? args.find((a) => !a.startsWith('--'))
		: args[flag].startsWith('--root=')
			? args[flag].slice(7)
			: args[flag + 1];
if (!rootArg)
	process.stderr.write('docx-viewer-mcp: no --root given; using the current working directory\n');

const server = createServer({ rootDir: rootArg ?? process.cwd() });
await server.connect(new StdioServerTransport());
