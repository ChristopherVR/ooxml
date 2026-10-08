import { describe, it, expect, vi } from 'vitest';
import { uploadTeamsServerFile } from './server-file-storage';

describe('host Teams storage', () => {
	it('uploads only to the configured server with redirects disabled', async () => {
		const fetch = vi
			.fn<typeof globalThis.fetch>()
			.mockResolvedValue(new Response(JSON.stringify({ url: '/files/workspace/brief.docx' })));
		const config = {
			mode: 'server' as const,
			syncUrl: 'wss://teams.test/sync',
			token: 'secret',
			iceServers: [],
		};
		const file = Object.assign(new Blob(['document']), { name: 'brief.docx' });
		expect(await uploadTeamsServerFile(config, 'workspace', file, { fetch })).toEqual({
			url: 'https://teams.test/files/workspace/brief.docx',
		});
		expect(fetch).toHaveBeenCalledWith(
			'https://teams.test/files/workspace/brief.docx',
			expect.objectContaining({
				redirect: 'error',
				credentials: 'omit',
				headers: { Authorization: 'Bearer secret' },
			}),
		);
	});
	it('rejects an off-origin URL returned by the server', async () => {
		const fetch = vi
			.fn<typeof globalThis.fetch>()
			.mockResolvedValue(new Response(JSON.stringify({ url: 'https://other.test/file' })));
		await expect(
			uploadTeamsServerFile(
				{ mode: 'server', syncUrl: 'wss://teams.test', iceServers: [] },
				'ws',
				Object.assign(new Blob(), { name: 'x.docx' }),
				{ fetch },
			),
		).rejects.toThrow('invalid file URL');
	});
});
