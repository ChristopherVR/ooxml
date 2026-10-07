import type { TeamsServerConfig } from 'ooxml-core/teams';

/** One ICE server per line: `urls[,urls] [username credential]`. */
export function parseIceLines(text: string): unknown[] {
	return text
		.split('\n')
		.map((line) => line.trim())
		.filter(Boolean)
		.map((line) => {
			const [urls = '', username, credential] = line.split(/\s+/u);
			return {
				urls: urls.split(',').filter(Boolean),
				...(username ? { username } : {}),
				...(credential ? { credential } : {}),
			};
		});
}

export const iceToText = (config: TeamsServerConfig): string =>
	config.iceServers
		.map((server) =>
			[[server.urls].flat().join(','), server.username, server.credential]
				.filter(Boolean)
				.join(' '),
		)
		.join('\n');
