import type { TeamsServerConfig } from './server-config';

/** The Teams server's file-upload protocol, shared by the default client and storage hosts. */
export async function uploadTeamsServerFile(
	config: TeamsServerConfig,
	workspaceId: string,
	file: Blob & { name: string },
	options: { signal?: AbortSignal; fetch?: typeof fetch } = {},
): Promise<{ url: string }> {
	if (config.mode !== 'server' || !config.syncUrl)
		throw new Error('Configure file storage before sharing attachments');
	const target = new URL(config.syncUrl);
	target.protocol = target.protocol === 'wss:' ? 'https:' : 'http:';
	const response = await (options.fetch ?? globalThis.fetch)(
		`${target.origin}/files/${encodeURIComponent(workspaceId)}/${encodeURIComponent(file.name)}`,
		{
			method: 'POST',
			body: file,
			redirect: 'error',
			credentials: 'omit',
			...(options.signal ? { signal: options.signal } : {}),
			...(config.token ? { headers: { Authorization: `Bearer ${config.token}` } } : {}),
		},
	);
	if (!response.ok) throw new Error(`Upload failed (${response.status})`);
	const result = (await response.json()) as { url?: unknown };
	if (
		typeof result.url !== 'string' ||
		!result.url.startsWith('/files/') ||
		result.url.includes('\\')
	)
		throw new Error('Server returned an invalid file URL');
	return { url: new URL(result.url, target.origin).href };
}
