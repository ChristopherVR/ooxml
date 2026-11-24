import { describe, expect, it } from 'vitest';
import { createCalcEngine } from './engine';
import { book, get } from './test-helpers';
import {
	createFetchImageProvider,
	createFetchWebserviceProvider,
	type FetchLike,
} from './web-providers';

const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('createFetchWebserviceProvider', () => {
	it('returns #BUSY! while fetching, caches the response, and notifies onUpdate', async () => {
		const updates: string[] = [];
		const fetcher: FetchLike = async (url) =>
			new Response(`<price>${url.endsWith('/ok') ? '42' : '0'}</price>`, { status: 200 });
		const webservice = createFetchWebserviceProvider({
			fetch: fetcher,
			onUpdate: (url) => updates.push(url),
		});
		const wb = book({ A1: '=WEBSERVICE("https://example.com/ok")' });
		const engine = createCalcEngine(wb, { providers: { webservice } });
		engine.recalculateAll();
		expect(get(wb, 'A1')).toEqual({ error: '#BUSY!' });

		await flush();
		expect(updates).toEqual(['https://example.com/ok']);
		engine.recalculateAll();
		expect(get(wb, 'A1')).toBe('<price>42</price>');
	});

	it('enforces allowedOrigins, handles network failures gracefully, and supports seed/clear', async () => {
		const failing: FetchLike = async () => {
			throw new Error('offline');
		};
		const webservice = createFetchWebserviceProvider({
			fetch: failing,
			allowedOrigins: ['https://allowed.example'],
		});
		expect(webservice.fetch('https://blocked.example/a')).toEqual({ error: '#BLOCKED!' });
		expect(webservice.fetch('https://allowed.example/fail')).toEqual({ error: '#BUSY!' });
		await flush();
		expect(webservice.fetch('https://allowed.example/fail')).toEqual({ error: '#CONNECT!' });

		webservice.seed('https://allowed.example/seeded', 'hello');
		expect(webservice.fetch('https://allowed.example/seeded')).toBe('hello');
		webservice.clear();
		expect(webservice.pending()).toBe(0);
	});
});

describe('createFetchImageProvider', () => {
	it('returns #BUSY! while fetching an image, caches bytes, and resolves alt text', async () => {
		const updates: string[] = [];
		const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);
		const fetcher: FetchLike = async () =>
			new Response(png, { status: 200, headers: { 'content-type': 'image/png' } });
		const image = createFetchImageProvider({
			fetch: fetcher,
			onUpdate: (url) => updates.push(url),
		});
		const wb = book({ A1: '=IMAGE("https://example.com/logo.png","Logo")' });
		const engine = createCalcEngine(wb, { providers: { image } });
		engine.recalculateAll();
		expect(get(wb, 'A1')).toEqual({ error: '#BUSY!' });

		await flush();
		expect(updates).toEqual(['https://example.com/logo.png']);
		engine.recalculateAll();
		expect(get(wb, 'A1')).toBe('Logo');
		expect(image.getImage('https://example.com/logo.png')?.mimeType).toBe('image/png');
		expect(image.getImage('https://example.com/logo.png')?.bytes).toEqual(png);
	});

	it('rejects non-image responses and disallowed origins', async () => {
		const fetcher: FetchLike = async () =>
			new Response('not an image', { status: 200, headers: { 'content-type': 'text/html' } });
		const image = createFetchImageProvider({
			fetch: fetcher,
			allowedOrigins: ['https://cdn.example'],
		});
		expect(
			image.resolve({ source: 'https://other.example/a.png', altText: '', sizing: 0 }),
		).toEqual({ error: '#BLOCKED!' });
		expect(image.resolve({ source: 'https://cdn.example/page', altText: '', sizing: 0 })).toEqual({
			error: '#BUSY!',
		});
		await flush();
		expect(image.resolve({ source: 'https://cdn.example/page', altText: '', sizing: 0 })).toEqual({
			error: '#VALUE!',
		});
	});
});
