#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './index.js';

// `--root <dir>` and `--formats a,b` (also `--name=value`); the legacy bare
// `<root> [formats]` arguments still work.
const positional = [];
const flags = {};
const args = process.argv.slice(2);
for (let i = 0; i < args.length; i++) {
	const arg = args[i];
	if (!arg.startsWith('--')) positional.push(arg);
	else if (arg.includes('='))
		flags[arg.slice(2, arg.indexOf('='))] = arg.slice(arg.indexOf('=') + 1);
	else flags[arg.slice(2)] = args[++i];
}
const root = flags.root ?? positional[0];
if (!root)
	process.stderr.write('ooxml-mcp: no --root given; using the current working directory\n');
const rootDir = root ?? process.cwd();
const selected = flags.formats ?? (flags.root ? positional[0] : positional[1]);
const options = selected ? { rootDir, formats: selected.split(',') } : { rootDir };
const server = await createServer(options);
await server.connect(new StdioServerTransport());
