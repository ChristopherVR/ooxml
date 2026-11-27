import type { OnlineVideoSearchProvider, OnlineVideoSearchResult } from './online-video-search';

export interface YouTubeVideoSearchOptions {
	/** The host's YouTube Data API key (the host's quota and terms; never bundled). */
	apiKey: string;
	/** Embed through `www.youtube-nocookie.com` (the default) or `www.youtube.com`. */
	privacyEnhanced?: boolean;
	/** Restrict results (`strict`, `moderate`, `none`); default `moderate`. */
	safeSearch?: 'strict' | 'moderate' | 'none';
	fetch?: typeof fetch;
}

interface SearchItem {
	id?: { videoId?: unknown };
	snippet?: {
		title?: unknown;
		channelTitle?: unknown;
		thumbnails?: Record<string, { url?: unknown } | undefined>;
	};
}

const textOf = (value: unknown) => (typeof value === 'string' ? value : undefined);

const ENTITIES: Record<string, string> = { amp: '&', quot: '"', apos: "'", lt: '<', gt: '>' };

/** The API returns titles HTML-escaped; they are shown as text. */
const unescapeHtml = (text: string) =>
	text.replace(/&(#\d{1,6}|[a-z]{2,4});/giu, (match, name: string) => {
		if (name.startsWith('#')) {
			const code = Number(name.slice(1));
			return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : match;
		}
		return ENTITIES[name.toLowerCase()] ?? match;
	});

/**
 * An adapter for the documented YouTube Data API v3 `search.list` endpoint. Not used unless the
 * host creates it and assigns it to `onlineVideoSearch`; add the embed host it returns
 * (`www.youtube-nocookie.com` or `www.youtube.com`) to `onlineVideoHosts`.
 */
export function createYouTubeVideoSearch(
	options: YouTubeVideoSearchOptions,
): OnlineVideoSearchProvider {
	const embedHost =
		options.privacyEnhanced === false ? 'www.youtube.com' : 'www.youtube-nocookie.com';
	return {
		name: 'YouTube',
		async search(query, { signal, locale }) {
			const url = new URL('https://www.googleapis.com/youtube/v3/search');
			url.search = new URLSearchParams({
				part: 'snippet',
				type: 'video',
				videoEmbeddable: 'true',
				maxResults: '20',
				safeSearch: options.safeSearch ?? 'moderate',
				relevanceLanguage: locale.split('-')[0] ?? 'en',
				q: query,
				key: options.apiKey,
			}).toString();
			const response = await (options.fetch ?? fetch)(url, {
				signal,
				credentials: 'omit',
				referrerPolicy: 'no-referrer',
			});
			if (!response.ok) throw new Error(`YouTube search failed (${response.status}).`);
			const body = (await response.json()) as { items?: SearchItem[] };
			const results: OnlineVideoSearchResult[] = [];
			for (const item of Array.isArray(body.items) ? body.items : []) {
				const id = textOf(item.id?.videoId);
				const title = textOf(item.snippet?.title);
				if (!id || !/^[\w-]{6,20}$/u.test(id) || !title) continue;
				const thumbs = item.snippet?.thumbnails ?? {};
				const thumbnail = textOf((thumbs.medium ?? thumbs.default)?.url);
				const publisher = textOf(item.snippet?.channelTitle);
				results.push({
					title: unescapeHtml(title),
					url: `https://${embedHost}/embed/${id}`,
					...(thumbnail ? { thumbnailUrl: thumbnail } : {}),
					...(publisher ? { publisher: unescapeHtml(publisher) } : {}),
				});
			}
			return results;
		},
	};
}
