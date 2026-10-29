import type { TeamsServerConfig } from './server-config';

/**
 * Headers that tell the server who is asking for a file: the signed session when there is one,
 * else (servers without accounts) the claimed user id. Shared by upload and link requests.
 */
export function fileRequestHeaders(
	config: TeamsServerConfig,
	userId?: string,
): Record<string, string> {
	return {
		...(config.token ? { Authorization: `Bearer ${config.token}` } : {}),
		...(config.sessionToken ? { 'X-Teams-Session': config.sessionToken } : {}),
		...(!config.sessionToken && userId ? { 'X-Teams-User': userId } : {}),
	};
}

/**
 * The Teams server's file-upload protocol, shared by the default client and storage hosts.
 * `scopeId` is the workspace id for channel files, or the chat room id for a private chat, which
 * the server restricts to that chat's members.
 */
export async function uploadTeamsServerFile(
	config: TeamsServerConfig,
	workspaceId: string,
	file: Blob & { name: string },
	options: { signal?: AbortSignal; fetch?: typeof fetch; userId?: string } = {},
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
			headers: fileRequestHeaders(config, options.userId),
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
