import { defineTeamsContentPreview, type TeamsContentPreview } from './content-preview.js';
import type { OpenFileDetail } from './teams-app.js';

beforeAll(() => defineTeamsContentPreview());
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

async function preview(detail: OpenFileDetail): Promise<TeamsContentPreview> {
	const element = document.createElement('teams-content-preview') as TeamsContentPreview;
	element.detail = detail;
	document.body.append(element);
	await element.updateComplete;
	return element;
}

describe('content preview', () => {
	it('renders Markdown as safe text and handles loading failure', async () => {
		vi.stubGlobal(
			'fetch',
			vi.fn(async () => new Response('# Title\n<script>alert(1)</script>\n```\n<b>code</b>\n```')),
		);
		const el = await preview({
			attachment: { name: 'notes.md', kind: 'other' },
			url: 'https://host/notes.md',
		});
		await vi.waitFor(() => expect(el.status).toBe('ready'));
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('[role=heading]')?.textContent?.trim()).toBe('Title');
		expect(el.shadowRoot!.querySelector('script')).toBeNull();
		expect(el.shadowRoot!.querySelector('code')?.textContent).toBe('<b>code</b>');
		el.detail = { attachment: { name: 'notes.md', kind: 'other' }, url: 'javascript:alert(1)' };
		await el.updateComplete;
		await vi.waitFor(() => expect(el.status).toBe('error'));
		expect(el.shadowRoot!.querySelector('a')).toBeNull();
	});

	it('sandboxes sites without same-origin or top navigation permissions', async () => {
		const el = await preview({
			attachment: { name: 'index.html', kind: 'other' },
			url: 'https://site.test/',
		});
		await el.updateComplete;
		const frame = el.shadowRoot!.querySelector('iframe')!;
		expect(frame.src).toBe('https://site.test/');
		expect(frame.getAttribute('sandbox')).toBe('allow-scripts allow-forms');
		expect(frame.getAttribute('referrerpolicy')).toBe('no-referrer');
	});

	it('supports host Office embeds after a native PowerPoint load failure', async () => {
		vi.stubGlobal(
			'fetch',
			vi.fn(async () => new Response('', { status: 404 })),
		);
		const el = await preview({
			attachment: { name: 'deck.pptx', kind: 'pptx' },
			url: 'https://files.test/deck.pptx?sig=1',
		});
		await vi.waitFor(() => expect(el.status).toBe('error'));
		el.embeds = { pptx: ({ url }) => `https://viewer.test/?source=${encodeURIComponent(url!)}` };
		await el.updateComplete;
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('iframe')?.src).toContain(
			'source=https%3A%2F%2Ffiles.test',
		);
	});

	it('aborts replaced content and ignores a late response', async () => {
		let resolve!: (response: Response) => void;
		const fetcher = vi.fn(
			(_url: string, _init: RequestInit) =>
				new Promise<Response>((r) => {
					resolve = r;
				}),
		);
		vi.stubGlobal('fetch', fetcher);
		const el = await preview({
			attachment: { name: 'old.md', kind: 'other' },
			url: 'https://host/old.md',
		});
		el.detail = { attachment: { name: 'new.html', kind: 'other' }, url: 'https://host/new.html' };
		await el.updateComplete;
		expect(fetcher.mock.calls[0]?.[1].signal?.aborted).toBe(true);
		resolve(new Response('# Old document'));
		await new Promise((r) => setTimeout(r, 10));
		expect(el.text).toBe('');
		expect(el.frame).toBe('https://host/new.html');
	});
});
