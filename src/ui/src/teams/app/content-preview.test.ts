import { defineTeamsContentPreview, type TeamsContentPreview } from './content-preview';
import type { OpenFileDetail } from './teams-app';

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
	it('renders accessible tables and read-only tasks with safe relative links', async () => {
		vi.stubGlobal(
			'fetch',
			vi.fn(
				async () =>
					new Response(
						'| File | Status |\n| :--- | ---: |\n| [Budget](./Budget.xlsx) | **Ready** |\n| <img src=x onerror=alert(1)> | [bad](javascript:x) |\n\n- [x] Reviewed\n- [ ] Waiting',
					),
			),
		);
		const el = await preview({
			attachment: { name: 'notes.md', kind: 'other' },
			url: 'https://files.test/project/notes.md?sig=secret',
		});
		await vi.waitFor(() => expect(el.status).toBe('ready'));
		await el.updateComplete;
		expect(el.shadowRoot!.querySelectorAll('th[scope=col]')).toHaveLength(2);
		expect(el.shadowRoot!.querySelector('td[data-align=right] strong')?.textContent).toBe('Ready');
		expect(el.shadowRoot!.querySelector('table a')?.getAttribute('href')).toBe(
			'https://files.test/project/Budget.xlsx',
		);
		expect(el.shadowRoot!.querySelector('table img')).toBeNull();
		expect(el.shadowRoot!.querySelectorAll('table a')).toHaveLength(1);
		const checks = Array.from(
			el.shadowRoot!.querySelectorAll<HTMLInputElement>('input[type=checkbox]'),
		);
		expect(
			checks.map((input) => [input.checked, input.disabled, input.getAttribute('aria-label')]),
		).toEqual([
			[true, true, 'Reviewed'],
			[false, true, 'Waiting'],
		]);
	});
	it('cancels workbook serialization without uploading or clearing local edits', async () => {
		vi.stubGlobal(
			'fetch',
			vi.fn(() => new Promise<Response>(() => {})),
		);
		const el = await preview({
			attachment: { name: 'Budget.xlsx', kind: 'xlsx' },
			url: 'https://files.test/Budget.xlsx',
		});
		const share = vi.fn(async () => {});
		el.saveCopy = share;
		el.native = 'xlsx';
		el.status = 'ready';
		el.editing = true;
		el.dirty = true;
		await el.updateComplete;
		let finish!: (bytes: Uint8Array) => void;
		const clean = vi.fn();
		Object.assign(el.shadowRoot!.querySelector('xlsx-editor')!, {
			commitEdit: () => true,
			saveBytes: () =>
				new Promise<Uint8Array>((resolve) => {
					finish = resolve;
				}),
			markClean: clean,
		});
		const button = (name: string) =>
			Array.from(el.shadowRoot!.querySelectorAll('button')).find(
				(entry) => entry.textContent?.trim() === name,
			)!;
		button('Save copy to channel').click();
		await el.updateComplete;
		button('Cancel workbook save').click();
		finish(new Uint8Array([1, 2, 3]));
		await el.updateComplete;
		expect(el.saving).toBe(false);
		expect(el.dirty).toBe(true);
		expect(el.saveError).toContain('Your edits remain local');
		expect(share).not.toHaveBeenCalled();
		expect(clean).not.toHaveBeenCalled();
	});
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
