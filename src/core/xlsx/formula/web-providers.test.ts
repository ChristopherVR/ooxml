import { describe, expect, it, vi } from 'vitest';
import { encodePng } from '../../drawingml/raster-encode';
import { createEditSession } from '../edit/session';
import { cellView } from '../layout/cell-view';
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
	it('returns #BUSY! while fetching and exposes picture bytes and per-cell alt text', async () => {
		const updates: string[] = [];
		const png = encodePng(2, 1, new Uint8Array(6));
		const fetcher: FetchLike = async () =>
			new Response(new Uint8Array(png), { status: 200, headers: { 'content-type': 'image/png' } });
		const image = createFetchImageProvider({
			fetch: fetcher,
			onUpdate: (url) => updates.push(url),
		});
		const wb = book({
			A1: '=IMAGE("https://example.com/logo.png","Logo")',
			A2: '=IMAGE("https://example.com/logo.png","Different",3,12)',
		});
		const engine = createCalcEngine(wb, { providers: { image } });
		engine.recalculateAll();
		expect(get(wb, 'A1')).toEqual({ error: '#BUSY!' });

		await flush();
		expect(updates).toEqual(['https://example.com/logo.png']);
		engine.recalculateAll();
		expect(cellView(wb, 0, 0, 0).picture).toMatchObject({
			altText: 'Logo',
			sizing: 0,
			pixelWidth: 2,
		});
		expect(cellView(wb, 0, 1, 0).picture).toMatchObject({
			altText: 'Different',
			sizing: 3,
			height: 12,
		});
		expect(image.getImage('https://example.com/logo.png')?.mimeType).toBe('image/png');
		expect(image.getImage('https://example.com/logo.png')?.bytes).toEqual(png);
	});

	it('restores the resolved image from the undo snapshot without recalculating', () => {
		const source = 'https://cdn.example/logo.png';
		const bytes = encodePng(2, 1, new Uint8Array(6));
		const image = createFetchImageProvider();
		image.seed(source, { url: source, bytes, mimeType: 'image/png', altText: '', sizing: 0 });
		const wb = book({ A1: `=IMAGE("${source}","Undo logo")` });
		createCalcEngine(wb, { providers: { image } }).recalculateAll();
		expect(cellView(wb, 0, 0, 0).picture?.altText).toBe('Undo logo');

		const session = createEditSession(wb, { recalc: false });
		session.setCellValue(0, 0, 0, 42);
		expect(cellView(wb, 0, 0, 0).picture).toBeUndefined();
		session.undo();
		expect(cellView(wb, 0, 0, 0).picture).toMatchObject({ altText: 'Undo logo', bytes });
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

	it('blocks HTTP even when a host fetcher is configured', () => {
		const fetch = vi.fn<FetchLike>();
		const provider = createFetchImageProvider({ fetch });
		expect(
			provider.resolve({ source: 'http://cdn.example/picture', altText: '', sizing: 0 }),
		).toEqual({ error: '#BLOCKED!' });
		expect(fetch).not.toHaveBeenCalled();
	});

	it('requires an explicit fetcher and cancels an oversized streamed response', async () => {
		const request = { source: 'https://cdn.example/picture', altText: '', sizing: 0 } as const;
		expect(createFetchImageProvider().resolve(request)).toEqual({ error: '#BLOCKED!' });
		let cancelled = false;
		const image = createFetchImageProvider({
			maxBytes: 4,
			fetch: async (url, init) => {
				expect(url).toBe(request.source);
				expect(init).toMatchObject({
					credentials: 'omit',
					redirect: 'error',
					referrerPolicy: 'no-referrer',
				});
				return new Response(
					new ReadableStream({
						pull: (controller) => controller.enqueue(new Uint8Array(3)),
						cancel: () => {
							cancelled = true;
						},
					}),
				);
			},
		});
		expect(image.resolve(request)).toEqual({ error: '#BUSY!' });
		await flush();
		expect(cancelled).toBe(true);
		expect(image.resolve(request)).toEqual({ error: '#BLOCKED!' });
	});
});
