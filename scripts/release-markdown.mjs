/**
 * release-markdown.mjs: the Markdown subset the release posts in `docs/releases/` are written in,
 * rendered to HTML with no dependencies. It is deliberately small; `build-release-notes.mjs`
 * documents the contract and turns the result into a page.
 *
 * Blocks: ATX headings (GitHub slugs), paragraphs, thematic breaks, bullet and numbered lists (nested by indent),
 * fenced code, tables, blockquotes and GitHub alerts (`> [!NOTE]`), HTML comments (dropped) and
 * raw HTML blocks (a line that starts with a tag, passed through to the next blank line). Inline:
 * bold, italics, code, links, images, autolinks and backslash escapes. Everything else is escaped.
 */

const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
export const escapeHtml = (text) => String(text).replace(/[&<>"']/g, (c) => ESCAPES[c]);

/** GitHub's heading slug: lower case, punctuation dropped, each space a hyphen. */
export function slug(text) {
	return text
		.toLowerCase()
		.replace(/[^\p{L}\p{M}\p{N}\p{Pc} -]/gu, '')
		.replace(/ /g, '-');
}

/** A slugger that numbers repeats the way GitHub does (`a`, `a-1`, `a-2`). */
export function slugger() {
	const seen = new Map();
	return (text) => {
		const base = slug(text);
		const count = seen.get(base) ?? 0;
		seen.set(base, count + 1);
		return count ? `${base}-${count}` : base;
	};
}

/** The text of inline Markdown without its markup, for slugs, titles and descriptions. */
export function plainText(md) {
	return md
		.replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1')
		.replace(/\[([^\]]*)\]\([^)]*\)/g, '$1')
		.replace(/`([^`]*)`/g, '$1')
		.replace(/(\*\*|__)(.+?)\1/g, '$2')
		.replace(/(^|[^\w*])[*_](?!\s)(.+?)[*_](?![\w*])/g, '$1$2')
		.replace(/\\([\\`*_{}[\]()#+\-.!|>~<])/g, '$1')
		.trim();
}

const SAFE_URL = /^(?:https?:|mailto:|#|\.{0,2}\/|[\w.-]+(?:[/#?]|$))/i;
const safeUrl = (url) => (SAFE_URL.test(url) && !/^\s*javascript:/i.test(url) ? url : '#');

/** Inline Markdown to HTML. `ctx.link(url, kind)` may rewrite link (`a`) and image (`img`) URLs. */
export function renderInline(md, ctx = {}) {
	const link = (url, kind) => escapeHtml(safeUrl(ctx.link ? ctx.link(url, kind) : url));
	let out = '';
	let i = 0;
	while (i < md.length) {
		const rest = md.slice(i);
		const prev = i ? md[i - 1] : ' ';
		let m;
		if ((m = /^\\([\\`*_{}[\]()#+\-.!|>~<])/.exec(rest))) out += escapeHtml(m[1]);
		else if ((m = /^(`+)([\s\S]*?[^`])\1(?!`)/.exec(rest)))
			out += `<code>${escapeHtml(m[2].replace(/^ (.*) $/, '$1'))}</code>`;
		else if ((m = /^!\[([^\]]*)\]\(\s*([^\s)]+)(?:\s+"([^"]*)")?\s*\)/.exec(rest)))
			out += `<img src="${link(m[2], 'img')}" alt="${escapeHtml(m[1])}"${m[3] ? ` title="${escapeHtml(m[3])}"` : ''} loading="lazy" />`;
		else if ((m = /^\[((?:\\.|`[^`]*`|[^\]])+)\]\(\s*([^\s)]+)(?:\s+"([^"]*)")?\s*\)/.exec(rest))) {
			const href = link(m[2], 'a');
			const external = /^https?:/i.test(href) ? ' rel="noopener"' : '';
			out += `<a href="${href}"${external}>${renderInline(m[1], ctx)}</a>`;
		} else if ((m = /^<(https?:\/\/[^\s<>]+)>/.exec(rest)))
			out += `<a href="${link(m[1], 'a')}" rel="noopener">${escapeHtml(m[1])}</a>`;
		else if ((m = /^(\*\*|__)(?!\s)([\s\S]+?)(?<!\s)\1/.exec(rest)))
			out += `<strong>${renderInline(m[2], ctx)}</strong>`;
		else if ((m = /^\*(?![\s*])([\s\S]+?)(?<![\s*])\*(?!\*)/.exec(rest)))
			out += `<em>${renderInline(m[1], ctx)}</em>`;
		else if (!/\w/.test(prev) && (m = /^_(?![\s_])([\s\S]+?)(?<![\s_])_(?!\w)/.exec(rest)))
			out += `<em>${renderInline(m[1], ctx)}</em>`;
		else {
			out += escapeHtml(md[i]);
			i += 1;
			continue;
		}
		i += m[0].length;
	}
	return out;
}

const FENCE = /^(\s*)(`{3,}|~{3,})\s*([\w+#.-]*)[^`]*$/;
const HEADING = /^ {0,3}(#{1,6})\s+(.*?)(?:\s+#+)?\s*$/;
const LIST = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/;
const HTML_BLOCK = /^ {0,3}<\/?[A-Za-z][\w-]*(?:\s|\/?>|$)/;
const QUOTE = /^ {0,3}>\s?/;
const TABLE_RULE = /^\s*\|?\s*:?-+:?\s*(?:\|\s*:?-+:?\s*)*\|?\s*$/;
const RULE = /^ {0,3}([-*_])(?:\s*\1){2,}\s*$/;
const ALERT = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*$/i;

const startsBlock = (line, next = '') =>
	FENCE.test(line) ||
	HEADING.test(line) ||
	LIST.test(line) ||
	QUOTE.test(line) ||
	HTML_BLOCK.test(line) ||
	line.trimStart().startsWith('<!--') ||
	(line.includes('|') && TABLE_RULE.test(next));

const cells = (row) =>
	row
		.trim()
		.replace(/^\|/, '')
		.replace(/(?<!\\)\|$/, '')
		.split(/(?<!\\)\|/)
		.map((cell) => cell.trim().replace(/\\\|/g, '|'));

/** Markdown source to a list of blocks: `{ type, ... }` objects that `renderBlocks` turns to HTML. */
export function parseBlocks(source) {
	const lines = source.replace(/\r\n?/g, '\n').split('\n');
	const blocks = [];
	let i = 0;
	while (i < lines.length) {
		const line = lines[i];
		if (!line.trim()) {
			i += 1;
			continue;
		}
		let m;
		if ((m = FENCE.exec(line))) {
			const [, indent, marker, lang] = m;
			const code = [];
			i += 1;
			while (i < lines.length && !lines[i].trim().startsWith(marker)) {
				code.push(lines[i].startsWith(indent) ? lines[i].slice(indent.length) : lines[i]);
				i += 1;
			}
			blocks.push({ type: 'code', lang: lang.toLowerCase(), text: code.join('\n') });
			i += 1;
		} else if (RULE.test(line)) {
			blocks.push({ type: 'rule' });
			i += 1;
		} else if ((m = HEADING.exec(line))) {
			blocks.push({ type: 'heading', level: m[1].length, text: m[2] });
			i += 1;
		} else if (line.trimStart().startsWith('<!--')) {
			while (i < lines.length && !lines[i].includes('-->')) i += 1;
			i += 1;
		} else if (HTML_BLOCK.test(line)) {
			const html = [];
			while (i < lines.length && lines[i].trim()) html.push(lines[i++]);
			blocks.push({ type: 'html', html: html.join('\n') });
		} else if (QUOTE.test(line)) {
			const inner = [];
			while (i < lines.length && QUOTE.test(lines[i])) inner.push(lines[i++].replace(QUOTE, ''));
			const alert = ALERT.exec(inner[0]?.trim() ?? '');
			if (alert) inner.shift();
			blocks.push({
				type: 'quote',
				alert: alert?.[1].toLowerCase(),
				blocks: parseBlocks(inner.join('\n')),
			});
		} else if (line.includes('|') && TABLE_RULE.test(lines[i + 1] ?? '')) {
			const head = cells(line);
			const align = cells(lines[i + 1]).map((c) =>
				c.endsWith(':')
					? c.startsWith(':')
						? 'center'
						: 'right'
					: c.startsWith(':')
						? 'left'
						: '',
			);
			const rows = [];
			i += 2;
			while (i < lines.length && lines[i].includes('|') && lines[i].trim())
				rows.push(cells(lines[i++]));
			blocks.push({ type: 'table', head, align, rows });
		} else if (LIST.test(line)) {
			const items = [];
			while (i < lines.length) {
				const current = lines[i];
				const item = LIST.exec(current);
				if (item)
					items.push({
						indent: item[1].replace(/\t/g, '    ').length,
						ordered: /\d/.test(item[2]),
						start: Number.parseInt(item[2], 10),
						text: item[3],
					});
				else if (current.trim() && !startsBlock(current) && items.length)
					items[items.length - 1].text += `\n${current.trim()}`;
				else if (!current.trim() && LIST.test(lines[i + 1] ?? '')) {
					i += 1;
					continue;
				} else break;
				i += 1;
			}
			blocks.push({ type: 'list', items });
		} else {
			const text = [line.trim()];
			i += 1;
			while (i < lines.length && lines[i].trim() && !startsBlock(lines[i], lines[i + 1]))
				text.push(lines[i++].trim());
			blocks.push({ type: 'paragraph', text: text.join('\n') });
		}
	}
	return blocks;
}

/** Flat list items (with indents) to nested `<ul>`/`<ol>` HTML. */
function renderList(items, ctx) {
	let html = '';
	const stack = [];
	const open = (item) => {
		const tag = item.ordered ? 'ol' : 'ul';
		const start = item.ordered && item.start !== 1 ? ` start="${item.start}"` : '';
		stack.push({ indent: item.indent, tag });
		return `<${tag}${start}>`;
	};
	for (const item of items) {
		while (stack.length > 1 && item.indent < stack[stack.length - 1].indent)
			html += `</li></${stack.pop().tag}>`;
		const top = stack[stack.length - 1];
		if (!top) html += open(item);
		else if (item.indent > top.indent) html += open(item);
		else html += '</li>';
		html += `<li>${renderInline(item.text.replace(/\n/g, ' '), ctx)}`;
	}
	while (stack.length) html += `</li></${stack.pop().tag}>`;
	return html;
}

const ALERT_LABELS = {
	note: 'Note',
	tip: 'Tip',
	important: 'Important',
	warning: 'Warning',
	caution: 'Caution',
};

/**
 * Blocks to HTML. `ctx.slug(text)` names headings, `ctx.code(block)`, `ctx.html(html)` and
 * `ctx.heading(block, id, html)` may replace the default rendering of those blocks.
 */
export function renderBlocks(blocks, ctx = {}) {
	const nextSlug = ctx.slug ?? slugger();
	const parts = [];
	for (const block of blocks) {
		if (block.type === 'heading') {
			const id = nextSlug(plainText(block.text));
			const inner = renderInline(block.text, ctx);
			parts.push(
				ctx.heading?.(block, id, inner) ??
					`<h${block.level} id="${escapeHtml(id)}">${inner}</h${block.level}>`,
			);
		} else if (block.type === 'paragraph')
			parts.push(`<p>${renderInline(block.text, ctx).replace(/\n/g, ' ')}</p>`);
		else if (block.type === 'list') parts.push(renderList(block.items, ctx));
		else if (block.type === 'rule') parts.push('<hr />');
		else if (block.type === 'code')
			parts.push(
				ctx.code?.(block) ??
					`<pre><code${block.lang ? ` class="language-${escapeHtml(block.lang)}"` : ''}>${escapeHtml(block.text)}</code></pre>`,
			);
		else if (block.type === 'html') parts.push(ctx.html ? ctx.html(block.html) : block.html);
		else if (block.type === 'quote') {
			const inner = renderBlocks(block.blocks, { ...ctx, slug: nextSlug });
			parts.push(
				block.alert
					? `<div class="callout callout-${block.alert}" role="note"><p class="callout-title">${ALERT_LABELS[block.alert]}</p>${inner}</div>`
					: `<blockquote>${inner}</blockquote>`,
			);
		} else if (block.type === 'table') {
			const cell = (tag, text, index) => {
				const align = block.align[index] ? ` style="text-align:${block.align[index]}"` : '';
				return `<${tag}${align}>${renderInline(text, ctx)}</${tag}>`;
			};
			const head = block.head.map((text, index) => cell('th', text, index)).join('');
			const rows = block.rows
				.map((row) => `<tr>${row.map((text, index) => cell('td', text, index)).join('')}</tr>`)
				.join('');
			parts.push(
				`<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>`,
			);
		}
	}
	return parts.join('\n');
}

/** Markdown to HTML in one call. */
export const renderMarkdown = (source, ctx) => renderBlocks(parseBlocks(source), ctx);
