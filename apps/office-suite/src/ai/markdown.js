/**
 * A small Markdown renderer for assistant replies. It builds DOM nodes and sets text through
 * `textContent`, so model output never becomes HTML; links are limited to http(s) and mailto.
 * It covers what chat replies use: headings, paragraphs, lists, quotes, tables, rules, fenced
 * code with a copy button, and inline code, bold, italic and links.
 */
const el = (tag, className) => {
	const node = document.createElement(tag);
	if (className) node.className = className;
	return node;
};

const inlinePattern =
	/(`+)([\s\S]*?)\1|\*\*([^*]+)\*\*|__([^_]+)__|\*([^*\s][^*]*)\*|_([^_\s][^_]*)_|\[([^\]]+)\]\(([^)\s]+)\)/gu;

function inline(target, text) {
	let last = 0;
	for (const match of text.matchAll(inlinePattern)) {
		if (match.index > last) target.append(text.slice(last, match.index));
		last = match.index + match[0].length;
		if (match[1]) target.append(Object.assign(el('code'), { textContent: match[2].trim() }));
		else if (match[3] ?? match[4]) inline(target.appendChild(el('strong')), match[3] ?? match[4]);
		else if (match[5] ?? match[6]) inline(target.appendChild(el('em')), match[5] ?? match[6]);
		else {
			let url;
			try {
				url = new URL(match[8]);
			} catch {}
			if (url && ['http:', 'https:', 'mailto:'].includes(url.protocol)) {
				const a = Object.assign(el('a'), {
					href: url.href,
					target: '_blank',
					rel: 'noopener noreferrer',
				});
				inline(a, match[7]);
				target.append(a);
			} else target.append(match[7]);
		}
	}
	if (last < text.length) target.append(text.slice(last));
}

function codeBlock(language, code) {
	const figure = el('figure', 'md-code');
	const bar = el('figcaption');
	const label = el('span');
	label.textContent = language || 'text';
	const copy = el('button');
	copy.type = 'button';
	copy.textContent = 'Copy';
	copy.onclick = () =>
		navigator.clipboard?.writeText(code).then(() => {
			copy.textContent = 'Copied';
			setTimeout(() => (copy.textContent = 'Copy'), 1500);
		});
	bar.append(label, copy);
	const pre = el('pre');
	pre.append(Object.assign(el('code'), { textContent: code }));
	figure.append(bar, pre);
	return figure;
}

const cells = (row) =>
	row
		.trim()
		.replace(/^\||\|$/gu, '')
		.split('|')
		.map((c) => c.trim());

function table(rows) {
	const wrap = el('div', 'md-table');
	const t = wrap.appendChild(el('table'));
	const head = t.appendChild(el('thead')).appendChild(el('tr'));
	for (const cell of cells(rows[0])) inline(head.appendChild(el('th')), cell);
	const body = t.appendChild(el('tbody'));
	for (const row of rows.slice(2)) {
		const tr = body.appendChild(el('tr'));
		for (const cell of cells(row)) inline(tr.appendChild(el('td')), cell);
	}
	return wrap;
}

/** Renders Markdown into a fragment of safe DOM nodes. */
export function renderMarkdown(source) {
	const out = document.createDocumentFragment();
	const lines = String(source).replace(/\r\n?/gu, '\n').split('\n');
	let i = 0;
	while (i < lines.length) {
		const line = lines[i];
		const fence = /^\s*(```+|~~~+)\s*([\w+-]*)/u.exec(line);
		if (fence) {
			const body = [];
			for (i++; i < lines.length && !lines[i].trim().startsWith(fence[1]); i++) body.push(lines[i]);
			i++;
			out.append(codeBlock(fence[2], body.join('\n')));
			continue;
		}
		if (!line.trim()) {
			i++;
			continue;
		}
		const heading = /^(#{1,6})\s+(.*)$/u.exec(line);
		if (heading) {
			inline(out.appendChild(el(`h${Math.min(6, heading[1].length + 2)}`)), heading[2]);
			i++;
			continue;
		}
		if (/^\s*([-*_])(\s*\1){2,}\s*$/u.test(line)) {
			out.append(el('hr'));
			i++;
			continue;
		}
		if (line.includes('|') && /^\s*\|?\s*:?-{3,}/u.test(lines[i + 1] ?? '')) {
			const rows = [];
			while (i < lines.length && lines[i].includes('|')) rows.push(lines[i++]);
			out.append(table(rows));
			continue;
		}
		if (/^\s*>/u.test(line)) {
			const quoted = [];
			while (i < lines.length && /^\s*>/u.test(lines[i]))
				quoted.push(lines[i++].replace(/^\s*>\s?/u, ''));
			out.appendChild(el('blockquote')).append(renderMarkdown(quoted.join('\n')));
			continue;
		}
		const item = /^\s*([-*+]|\d+[.)])\s+/u;
		if (item.test(line)) {
			const ordered = /^\s*\d/u.test(line);
			const list = out.appendChild(el(ordered ? 'ol' : 'ul'));
			while (i < lines.length && item.test(lines[i]) && /^\s*\d/u.test(lines[i]) === ordered) {
				const text = [lines[i++].replace(item, '')];
				while (i < lines.length && /^\s{2,}\S/u.test(lines[i]) && !item.test(lines[i]))
					text.push(lines[i++].trim());
				inline(list.appendChild(el('li')), text.join(' '));
			}
			continue;
		}
		const paragraph = [];
		while (
			i < lines.length &&
			lines[i].trim() &&
			!/^(#{1,6}\s|\s*```|\s*~~~|\s*>|\s*([-*+]|\d+[.)])\s)/u.test(lines[i])
		)
			paragraph.push(lines[i++]);
		const p = out.appendChild(el('p'));
		paragraph.forEach((text, index) => {
			if (index) p.append(el('br'));
			inline(p, text);
		});
	}
	return out;
}
