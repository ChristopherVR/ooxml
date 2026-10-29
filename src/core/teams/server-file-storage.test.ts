import { describe, it, expect, vi } from 'vitest';
import { fileRequestHeaders, uploadTeamsServerFile } from './server-file-storage';

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
	it('names the signed session, or the claimed user without one', async () => {
		const base = { mode: 'server' as const, iceServers: [] };
		expect(fileRequestHeaders({ ...base, token: 't', sessionToken: 's' }, 'ada')).toEqual({
			Authorization: 'Bearer t',
			'X-Teams-Session': 's',
		});
		expect(fileRequestHeaders(base, 'ada')).toEqual({ 'X-Teams-User': 'ada' });
		expect(fileRequestHeaders(base)).toEqual({});
		const fetch = vi
			.fn<typeof globalThis.fetch>()
			.mockResolvedValue(new Response(JSON.stringify({ url: '/files/chat-1/a.docx' })));
		await uploadTeamsServerFile(
			{ ...base, syncUrl: 'wss://teams.test/sync', sessionToken: 's' },
			'chat-1',
			Object.assign(new Blob(['x']), { name: 'a.docx' }),
			{ fetch },
		);
		expect(fetch.mock.calls[0]?.[0]).toBe('https://teams.test/files/chat-1/a.docx');
		expect(fetch.mock.calls[0]?.[1]?.headers).toEqual({ 'X-Teams-Session': 's' });
	});
});
