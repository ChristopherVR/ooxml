import { mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { describe, it, expect } from 'vitest';

import { createServer } from '../../mcp/server';

async function connect(rootDir: string) {
	const server = createServer({ rootDir });
	const client = new Client({ name: 'test-client', version: '1.0.0' });
	const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
	await server.connect(serverTransport);
	await client.connect(clientTransport);
	return { client, server };
}

const text = (result: { content: unknown }) =>
	(result.content as { type: string; text: string }[])[0]?.text ?? '';

describe('pptx MCP realpath confinement', () => {
	it('rejects a link inside the root that points outside it, without naming the root', async () => {
		const root = await mkdtemp(join(tmpdir(), 'pptx-root-'));
		const outside = await mkdtemp(join(tmpdir(), 'pptx-outside-'));
		const { client, server } = await connect(root);
		try {
			await writeFile(join(outside, 'secret.pptx'), new Uint8Array([1]));
			try {
				await symlink(
					outside,
					join(root, 'escape'),
					process.platform === 'win32' ? 'junction' : 'dir',
				);
			} catch {
				return;
			}
			const failure = await client.callTool({
				name: 'get_slide',
				arguments: { filePath: 'escape/secret.pptx', slideIndex: 0 },
			});
			expect(failure.isError).toBe(true);
			expect(text(failure)).toContain('outside');
			expect(text(failure)).not.toContain(root);
		} finally {
			await client.close();
			await server.close();
			await rm(root, { recursive: true, force: true });
			await rm(outside, { recursive: true, force: true });
		}
	});

	it('refuses a Markdown output directory outside the root', async () => {
		const root = await mkdtemp(join(tmpdir(), 'pptx-root-'));
		const { client, server } = await connect(root);
		try {
			const failure = await client.callTool({
				name: 'convert_to_markdown',
				arguments: { filePath: 'deck.pptx', outputDir: '../out' },
			});
			expect(failure.isError).toBe(true);
		} finally {
			await client.close();
			await server.close();
			await rm(root, { recursive: true, force: true });
		}
	});
});
