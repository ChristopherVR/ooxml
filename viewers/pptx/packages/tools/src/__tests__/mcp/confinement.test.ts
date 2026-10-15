import { mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { describe, it, expect } from 'vitest';

import { confineOutputDir, runMcpTool } from '../../mcp/handlers';

describe('pptx MCP realpath confinement', () => {
	it('rejects a link inside the root that points outside it', async () => {
		const root = await mkdtemp(join(tmpdir(), 'pptx-root-'));
		const outside = await mkdtemp(join(tmpdir(), 'pptx-outside-'));
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
			const failure = await runMcpTool(
				'escape/secret.pptx',
				() => {
					throw new Error('must not run');
				},
				{ rootDir: root },
			).catch((e: Error) => e);
			expect((failure as Error).message).toContain('outside');
			expect((failure as Error).message).not.toContain(root);
			await expect(confineOutputDir('escape/out', { rootDir: root })).rejects.toThrow('outside');
		} finally {
			await rm(root, { recursive: true, force: true });
			await rm(outside, { recursive: true, force: true });
		}
	});

	it('returns a root-relative output directory', async () => {
		const root = await mkdtemp(join(tmpdir(), 'pptx-root-'));
		try {
			expect(await confineOutputDir('out/media', { rootDir: root })).toBe('out/media');
			expect(await confineOutputDir('.', { rootDir: root })).toBe('.');
		} finally {
			await rm(root, { recursive: true, force: true });
		}
	});
});
