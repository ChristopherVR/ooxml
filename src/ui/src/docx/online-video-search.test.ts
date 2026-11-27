import { describe, expect, it, vi } from 'vitest';
import { allowedVideoResults, formatVideoDuration, MAX_VIDEO_RESULTS } from './online-video-search';
import { createYouTubeVideoSearch } from './online-video-youtube';

describe('online video search results', () => {
	it('lists only https results on an allowed host, once each', () => {
		const results = allowedVideoResults(
			[
				{
					title: 'Allowed',
					url: 'https://video.example.com/embed/1',
					thumbnailUrl: 'http://insecure.test/y.png',
				},
				{ title: 'Again', url: 'https://video.example.com/embed/1' },
				{ title: 'Other host', url: 'https://elsewhere.test/embed/2' },
				{ title: 'Plain http', url: 'http://video.example.com/embed/3' },
				{ title: 'Script', url: 'javascript:' },
			],
			['video.example.com'],
		);
		expect(results).toEqual([{ title: 'Allowed', url: 'https://video.example.com/embed/1' }]);
	});

	it('caps the list', () => {
		const many = Array.from({ length: MAX_VIDEO_RESULTS + 3 }, (_, i) => ({
			title: `V${i}`,
			url: `https://video.example.com/embed/${i}`,
		}));
		expect(allowedVideoResults(many, ['video.example.com'])).toHaveLength(MAX_VIDEO_RESULTS);
	});

	it('formats durations', () => {
		expect(formatVideoDuration(65)).toBe('1:05');
		expect(formatVideoDuration(3725)).toBe('1:02:05');
	});
});

describe('YouTube video search adapter', () => {
	it('sends only the query, language and key, and maps results to embed addresses', async () => {
		const fetch = vi.fn(async (url: URL | RequestInfo, _init?: RequestInit) => {
			const query = new URL(String(url)).searchParams;
			expect(query.get('q')).toBe('cats');
			expect(query.get('key')).toBe('host-key');
			expect(query.get('relevanceLanguage')).toBe('de');
			return new Response(
				JSON.stringify({
					items: [
						{
							id: { videoId: 'abcDEF12345' },
							snippet: {
								title: 'Tom &amp; Jerry&#39;s',
								channelTitle: 'Studio',
								thumbnails: { medium: { url: 'https://i.ytimg.com/vi/abc/mq.jpg' } },
							},
						},
						{ id: { videoId: '../bad' }, snippet: { title: 'Rejected' } },
					],
				}),
			);
		});
		const provider = createYouTubeVideoSearch({
			apiKey: 'host-key',
			fetch: fetch as unknown as typeof globalThis.fetch,
		});
		const results = await provider.search('cats', {
			signal: new AbortController().signal,
			locale: 'de-DE',
		});
		expect(results).toEqual([
			{
				title: "Tom & Jerry's",
				url: 'https://www.youtube-nocookie.com/embed/abcDEF12345',
				thumbnailUrl: 'https://i.ytimg.com/vi/abc/mq.jpg',
				publisher: 'Studio',
			},
		]);
		expect(fetch.mock.calls[0]?.[1]).toMatchObject({
			credentials: 'omit',
			referrerPolicy: 'no-referrer',
		});
	});
});
