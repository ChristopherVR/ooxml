import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { z } from 'zod';
import { createServer } from '../src/index.js';

test('combined view composes selected registrations and forwards their root', async () => {
	const seen = [];
	const registration = (format) => ({
		registerTools(server, options) {
			seen.push(options.rootDir);
			server.registerTool(
				format + '_example',
				{ inputSchema: { value: z.string() } },
				async (p) => ({
					content: [{ type: 'text', text: p.value }],
				}),
			);
		},
	});
	const server = await createServer({
		rootDir: '/documents',
		formats: ['docx', 'xlsx'],
		registrations: { docx: registration('docx'), xlsx: registration('xlsx') },
	});
	const client = new Client({ name: 'test', version: '1.0.0' });
	const [a, b] = InMemoryTransport.createLinkedPair();
	try {
		await server.connect(b);
		await client.connect(a);
		assert.deepEqual(seen, ['/documents', '/documents']);
		assert.deepEqual((await client.listTools()).tools.map((tool) => tool.name).sort(), [
			'docx_example',
			'xlsx_example',
		]);
		const result = await client.callTool({
			name: 'xlsx_example',
			arguments: { value: 'forwarded' },
		});
		assert.equal(result.content[0].text, 'forwarded');
	} finally {
		await client.close();
		await server.close();
	}
});

test('invalid selections and incompatible packages fail explicitly', async () => {
	await assert.rejects(createServer({ formats: [] }), /at least one/);
	await assert.rejects(createServer({ formats: ['docx', 'docx'] }), /duplicates/);
	await assert.rejects(createServer({ formats: ['unknown'] }), /Unknown format/);
	await assert.rejects(createServer({ formats: ['__proto__'] }), /Unknown format/);
	await assert.rejects(
		createServer({ formats: ['docx'], registrations: { docx: {} } }),
		/registerTools/,
	);
});
