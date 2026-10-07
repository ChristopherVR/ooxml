import { detectOfficeKind, type OfficeKind } from './model';
import { markdownTable, type MarkdownTable } from './markdown-table';
export type { MarkdownTable } from './markdown-table';

export type ContentKind = OfficeKind | 'markdown' | 'website' | 'text';

/** Extension takes precedence, including for non-Office content. */
export function detectContentKind(name: string, mime?: string): ContentKind {
	const ext = name.split('.').pop()?.toLowerCase();
	if (ext === 'md' || ext === 'markdown') return 'markdown';
	if (ext === 'html' || ext === 'htm') return 'website';
	if (ext === 'txt' || ext === 'csv' || ext === 'json') return 'text';
	const office = detectOfficeKind(name);
	if (office !== 'other') return office;
	const type = mime?.split(';')[0]?.trim().toLowerCase();
	if (type === 'text/markdown') return 'markdown';
	if (type === 'text/html') return 'website';
	if (type === 'text/plain') return 'text';
	return detectOfficeKind(name, mime);
}

/** Only web URLs and origin-relative paths, with no credentials or control characters. */
export function contentUrl(value: string | undefined, base: string): string | null {
	if (!value || /[\u0000-\u0020\u007f\\]/u.test(value)) return null;
	if (!/^(https?:\/\/|\/(?!\/))/iu.test(value)) return null;
	try {
		const url = new URL(value, base);
		return /^https?:$/u.test(url.protocol) && !url.username && !url.password ? url.href : null;
	} catch {
		return null;
	}
}

/** Bounded streamed reads, shared by every content renderer. No server token is forwarded. */
export async function readContent(
	url: string,
	signal: AbortSignal,
	maxBytes: number,
	doFetch: typeof fetch = globalThis.fetch,
): Promise<Uint8Array> {
	const response = await doFetch(url, {
		signal,
		credentials: 'omit',
		referrerPolicy: 'no-referrer',
	});
	if (!response.ok) throw new Error(`Could not load content (${response.status})`);
	const reader = response.body?.getReader();
	if (!reader) throw new Error('Content streaming is unavailable');
	const chunks: Uint8Array[] = [];
	let size = 0;
	try {
		if (Number(response.headers.get('content-length')) > maxBytes)
			throw new Error('Content is too large to preview');
		while (true) {
			const { value, done } = await reader.read();
			if (done) break;
			size += value.byteLength;
			if (size > maxBytes) throw new Error('Content is too large to preview');
			chunks.push(value);
		}
	} catch (error) {
		await reader.cancel().catch(() => {});
		throw error;
	} finally {
		reader.releaseLock();
	}
	const bytes = new Uint8Array(size);
	let offset = 0;
	for (const chunk of chunks) {
		bytes.set(chunk, offset);
		offset += chunk.byteLength;
	}
	return bytes;
}

export interface MarkdownBlock {
	kind: 'heading' | 'paragraph' | 'code' | 'quote' | 'list' | 'table';
	text: string;
	level: number;
	checked?: boolean;
	table?: MarkdownTable;
}

/** Resolve Markdown links against their file, without expanding the embedding URL policy. */
export function markdownUrl(value: string, base: string): string | null {
	if (!value || /[\u0000-\u0020\u007f\\]/u.test(value) || value.startsWith('//')) return null;
	if (/^[a-z][a-z\d+.-]*:/iu.test(value) && !/^https?:\/\//iu.test(value)) return null;
	if (!contentUrl(base, base)) return null;
	try {
		return contentUrl(new URL(value, base).href, base);
	} catch {
		return null;
	}
}

export interface MarkdownInline {
	kind: 'text' | 'strong' | 'emphasis' | 'code' | 'link';
	text: string;
	url?: string;
}

/** Flat inline subset; no embedded HTML, images or automatic resource requests. */
export function markdownInline(text: string, base: string): MarkdownInline[] {
	const tokens: MarkdownInline[] = [];
	const syntax = /`([^`]+)`|\*\*([^*]+)\*\*|\*([^*]+)\*|\[([^\]]+)\]\(([^\s)]+)\)/gu;
	let offset = 0;
	for (const match of text.matchAll(syntax)) {
		if (match.index > offset) tokens.push({ kind: 'text', text: text.slice(offset, match.index) });
		if (match[1]) tokens.push({ kind: 'code', text: match[1] });
		else if (match[2]) tokens.push({ kind: 'strong', text: match[2] });
		else if (match[3]) tokens.push({ kind: 'emphasis', text: match[3] });
		else {
			const url = markdownUrl(match[5]!, base);
			tokens.push(url ? { kind: 'link', text: match[4]!, url } : { kind: 'text', text: match[0] });
		}
		offset = match.index + match[0].length;
	}
	if (offset < text.length) tokens.push({ kind: 'text', text: text.slice(offset) });
	return tokens;
}

/** A deliberately small Markdown subset. Raw HTML stays text; never emits HTML strings. */
export function markdownBlocks(source: string): MarkdownBlock[] {
	const blocks: MarkdownBlock[] = [];
	const lines = source.replace(/\r\n?/gu, '\n').split('\n');
	let code: string[] | null = null;
	let fence = '';
	for (let index = 0; index < lines.length; index++) {
		const line = lines[index]!;
		const marker = /^\s{0,3}(`{3,}|~{3,})(.*)$/u.exec(line);
		if (code) {
			if (
				marker &&
				marker[1]![0] === fence[0] &&
				marker[1]!.length >= fence.length &&
				!marker[2]!.trim()
			) {
				blocks.push({ kind: 'code', text: code.join('\n'), level: 0 });
				code = null;
			} else code.push(line);
			continue;
		}
		if (marker) {
			code = [];
			fence = marker[1]!;
			continue;
		}
		if (!line.trim()) continue;
		const heading = /^(#{1,6})\s+(.+)$/u.exec(line);
		const list = /^\s*[-*+]\s+(.+)$/u.exec(line);
		const table = !heading && !list && !/^>\s?/u.test(line) ? markdownTable(lines, index) : null;
		if (table) {
			blocks.push({ kind: 'table', text: '', level: 0, table: table.table });
			index = table.end;
			continue;
		}
		if (heading) blocks.push({ kind: 'heading', text: heading[2]!, level: heading[1]!.length });
		else if (list) {
			const task = /^\[([ xX\t])\]\s+(.+)$/u.exec(list[1]!);
			blocks.push(
				task
					? { kind: 'list', text: task[2]!, level: 0, checked: /x/iu.test(task[1]!) }
					: { kind: 'list', text: list[1]!, level: 0 },
			);
		} else if (/^>\s?/u.test(line))
			blocks.push({ kind: 'quote', text: line.replace(/^>\s?/u, ''), level: 0 });
		else blocks.push({ kind: 'paragraph', text: line, level: 0 });
	}
	if (code) blocks.push({ kind: 'code', text: code.join('\n'), level: 0 });
	return blocks;
}
