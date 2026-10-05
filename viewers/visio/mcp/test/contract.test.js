import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { createServer } from '../src/index.js';

test('visio MCP discovery, input validation and tool errors', async () => {
	const root = await mkdtemp(join(tmpdir(), 'visio-mcp-'));
	const server = createServer({ rootDir: root });
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	try {
		await server.connect(b);
		await client.connect(a);
		const { tools } = await client.listTools();
		assert.ok(tools.length >= 2);
		assert.ok(tools.every((tool) => tool.name.startsWith('visio_')));
		const inspect = tools.find((tool) => tool.name === 'visio_inspect');
		assert.equal(inspect.annotations.readOnlyHint, true);
		const missing = await client.callTool({
			name: 'visio_inspect',
			arguments: { filePath: 'missing.vsdx' },
		});
		assert.equal(missing.isError, true);
		const invalid = await client.callTool({ name: 'visio_inspect', arguments: { filePath: 123 } });
		assert.equal(invalid.isError, true);
	} finally {
		await client.close();
		await server.close();
		await rm(root, { recursive: true, force: true });
	}
});
