#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './index.js';

const rootDir = process.argv[2] ?? process.cwd();
const selected = process.argv[3];
const options = selected ? { rootDir, formats: selected.split(',') } : { rootDir };
const server = await createServer(options);
await server.connect(new StdioServerTransport());
