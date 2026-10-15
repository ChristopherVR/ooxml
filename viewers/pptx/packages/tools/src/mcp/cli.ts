#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

import { createServer } from './server';

/** `--root <dir>` (or `--root=<dir>`); otherwise PPTX_TOOLS_ROOT, then the cwd with a warning. */
function parseRoot(argv: readonly string[]): string | undefined {
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i] ?? '';
		if (arg === '--root') return argv[i + 1];
		if (arg.startsWith('--root=')) return arg.slice('--root='.length);
	}
	return undefined;
}

const rootDir = parseRoot(process.argv.slice(2));
void createServer(rootDir ? { rootDir } : undefined).connect(new StdioServerTransport());
