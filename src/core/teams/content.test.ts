import {
	contentUrl,
	detectContentKind,
	markdownBlocks,
	markdownInline,
	markdownUrl,
	readContent,
} from './content.js';
import { filesOf } from './view.js';

describe('content routing', () => {
	it.each([
		['Plan.DOCX', 'text/html', 'docx'],
		['notes.MD', 'text/html', 'markdown'],
		['index.html', undefined, 'website'],
		['data.json', undefined, 'text'],
		['attachment', 'text/markdown; charset=utf-8', 'markdown'],
		['deck.pptx', undefined, 'pptx'],
		['book.xlsx', undefined, 'xlsx'],
		['drawing.vsdx', undefined, 'vsdx'],
		['unknown.bin', undefined, 'other'],
	])('routes %s', (name, mime, kind) => expect(detectContentKind(name, mime)).toBe(kind));

	it.each([
		'javascript:alert(1)',
		'data:text/html,x',
		'//evil.test',
		'/\\evil.test',
		'https://user:pass@host/file',
		'https://host/\nfile',
	])('rejects unsafe URL %s', (url) => {
		expect(contentUrl(url, 'https://workspace.test')).toBeNull();
	});
	it('keeps signed queries and resolves same-origin paths', () => {
		expect(contentUrl('/files/a?sig=secret&exp=1', 'https://workspace.test')).toBe(
			'https://workspace.test/files/a?sig=secret&exp=1',
		);
	});
	it('preserves MIME information in the Files view', () => {
		const [file] = filesOf([
			{
				id: 'm',
				channelId: 'c',
				authorId: 'a',
				authorName: 'Ada',
				text: '',
				ts: 1,
				deleted: false,
				reactions: {},
				attachments: [{ name: 'file', kind: 'other', mime: 'text/markdown', url: '/file' }],
			},
		]);
		expect(file?.mime).toBe('text/markdown');
	});
});

describe('Markdown blocks', () => {
	it('bounds sparse table expansion while preserving the remaining source', () => {
		const header = `|${Array(128).fill('A').join('|')}|`;
		const delimiter = `|${Array(128).fill('-').join('|')}|`;
		const blocks = markdownBlocks(`${header}\n${delimiter}\n${Array(200).fill('row').join('\n')}`);
		expect(blocks[0]?.table?.rows).toHaveLength(127);
		expect(blocks.filter((block) => block.kind === 'paragraph')).toHaveLength(73);
		expect(
			markdownBlocks(`${header.slice(0, -1)}|Extra|\n${delimiter.slice(0, -1)}|-|`)[0]?.kind,
		).toBe('paragraph');
	});
	it('parses table alignment, escaped pipes and uneven body rows', () => {
		const [block, quote] = markdownBlocks(
			'| Name | Value | Status |\n| :--- | ---: | :---: |\n| A\\|B | **12** | `x\\|y` |\nshort\n| Extra | 2 | Done | ignored |\n> after',
		);
		expect(block?.table).toEqual({
			headers: ['Name', 'Value', 'Status'],
			align: ['left', 'right', 'center'],
			rows: [
				['A|B', '**12**', '`x|y`'],
				['short', '', ''],
				['Extra', '2', 'Done'],
			],
		});
		expect(quote?.kind).toBe('quote');
	});
	it('leaves invalid table delimiters as text and does not parse tables in fences', () => {
		expect(markdownBlocks('| A | B |\n| --- |').some((block) => block.kind === 'table')).toBe(
			false,
		);
		expect(markdownBlocks('| A |\n| wrong |').some((block) => block.kind === 'table')).toBe(false);
		expect(markdownBlocks('```\n| A |\n| --- |\n```')[0]?.kind).toBe('code');
		expect(markdownBlocks('A\n---')[0]?.kind).not.toBe('table');
	});
	it('renders task state only for valid flat list markers', () => {
		expect(
			markdownBlocks(
				'- [x] Done\n- [X] Also done\n- [ ] Pending\n- [z] Literal\n[x] Paragraph',
			).map((block) => [block.text, block.checked]),
		).toEqual([
			['Done', true],
			['Also done', true],
			['Pending', false],
			['[z] Literal', undefined],
			['[x] Paragraph', undefined],
		]);
	});
	it('resolves file-relative links without forwarding signed queries or allowing unsafe schemes', () => {
		const base = 'https://files.test/project/notes.md?sig=secret';
		expect(markdownUrl('./Budget.xlsx', base)).toBe('https://files.test/project/Budget.xlsx');
		expect(markdownUrl('../docs/guide.md#review', base)).toBe(
			'https://files.test/docs/guide.md#review',
		);
		expect(markdownInline('[Budget](./Budget.xlsx)', base)[0]?.url).toBe(
			'https://files.test/project/Budget.xlsx',
		);
		for (const url of [
			'//evil.test/file',
			'javascript:x',
			'data:x',
			'mailto:x',
			'https://user:pass@host/file',
			'https:\\evil.test',
			'file\nname',
		])
			expect(markdownUrl(url, base)).toBeNull();
	});
	it('renders inline formatting and rejects unsafe links', () => {
		const tokens = markdownInline(
			'**bold** *italic* `code` [safe](https://host/a) [bad](javascript:x)',
			'https://host/notes.md',
		);
		expect(tokens.filter((t) => t.kind !== 'text')).toEqual([
			{ kind: 'strong', text: 'bold' },
			{ kind: 'emphasis', text: 'italic' },
			{ kind: 'code', text: 'code' },
			{ kind: 'link', text: 'safe', url: 'https://host/a' },
		]);
		expect(tokens.at(-1)).toEqual({ kind: 'text', text: '[bad](javascript:x)' });
	});
	it('supports headings, bullets, quotes, fences and leaves HTML as text', () => {
		expect(
			markdownBlocks('# Title\r\n- item\n> quote\n```js\n<img onerror=x>\n```\n<script>x</script>'),
		).toEqual([
			{ kind: 'heading', text: 'Title', level: 1 },
			{ kind: 'list', text: 'item', level: 0 },
			{ kind: 'quote', text: 'quote', level: 0 },
			{ kind: 'code', text: '<img onerror=x>', level: 0 },
			{ kind: 'paragraph', text: '<script>x</script>', level: 0 },
		]);
	});
	it('keeps an unclosed fence and does not close a fence with a different marker', () => {
		expect(markdownBlocks('~~~\n```\nx')[0]).toEqual({ kind: 'code', text: '```\nx', level: 0 });
	});
});

describe('content loading', () => {
	it('reads bounded content without credentials', async () => {
		const fetcher = vi.fn(
			async (_url: string | URL | Request, _options?: RequestInit) => new Response('hello'),
		);
		expect(
			new TextDecoder().decode(
				await readContent('https://host/file', new AbortController().signal, 10, fetcher),
			),
		).toBe('hello');
		expect(fetcher.mock.calls[0]?.[1]).toMatchObject({
			credentials: 'omit',
			referrerPolicy: 'no-referrer',
		});
	});
	it('rejects HTTP failures and oversized bodies without content-length', async () => {
		await expect(
			readContent('/x', new AbortController().signal, 3, async () => new Response('abcd')),
		).rejects.toThrow('too large');
		await expect(
			readContent(
				'/x',
				new AbortController().signal,
				10,
				async () => new Response('', { status: 403 }),
			),
		).rejects.toThrow('403');
	});
});
