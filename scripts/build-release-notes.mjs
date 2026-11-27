#!/usr/bin/env node
/**
 * build-release-notes.mjs: render each release post in `docs/releases/<name>.md` into the Pages
 * page `releases/<name>/index.html`, on the template and styles in `site/releases/`.
 *
 *   node scripts/build-release-notes.mjs [out]      (out defaults to suite-dist)
 *
 * `build-suite.mjs` and `build-pages.mjs` call `writeReleasePages(out)` after copying `site/`.
 * The post is a GitHub-readable Markdown file; the page reads these parts of it (see the
 * "Release post" section of docs/pages-design.md for the whole contract):
 *
 * - `# Title`, then a paragraph with `**Published:**` (the metadata line, ` | ` between items),
 *   then optionally an image alone in a paragraph: the hero.
 * - Everything before the first `##` is the lede. Its first bullet list whose items all open
 *   with bold text is drawn as the pillar cards.
 * - Each `##` is a section in the table of contents, with its `###` headings under it.
 * - `**file.ts**` alone on the line before a fence names the code block; `_Caption._` alone on
 *   the line after a fence or an image is its caption. A paragraph that starts with
 *   `Learn more` is the feature's closing link.
 * - `> [!NOTE]` (TIP, IMPORTANT, WARNING, CAUTION) is a callout.
 * - `<a id="changes-<area>"></a>` before a `<details>` block marks it as an area's complete list;
 *   its `<summary>` takes the area's tile and a count of the bullets inside.
 *
 * A post whose `showcase.json` names an `edition` gets a look of its own on the same template:
 * see release-edition.mjs.
 *
 * Relative links resolve against the post's folder: pictures are copied next to the page,
 * related release pages link locally and other repository paths link to GitHub.
 */
import {
	copyFileSync,
	existsSync,
	mkdirSync,
	readFileSync,
	readdirSync,
	rmSync,
	writeFileSync,
} from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

import {
	chapterProduct,
	chapterTile,
	editionAssets,
	editionOf,
	renderChapterLinks,
	renderEditionIndex,
	renderFeature,
} from './release-edition.mjs';
import { LANGUAGE_NAMES, highlight } from './release-highlight.mjs';
import { mediaResolver } from './release-media.mjs';
import { renderReleaseIndex, renderShowcase } from './release-showcase.mjs';
import {
	escapeHtml,
	parseBlocks,
	plainText,
	renderBlocks,
	renderInline,
	slug,
	slugger,
} from './release-markdown.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const POSTS_DIR = 'docs/releases';
export const TEMPLATE = 'site/releases/template.html';

/** The areas of the complete change lists, by the id after `changes-`, with their tiles. */
export const AREAS = {
	core: { tile: '', product: 'office' },
	'ui-mcp': { tile: 'U', product: 'office' },
	word: { tile: 'W', product: 'word' },
	excel: { tile: 'X', product: 'excel' },
	powerpoint: { tile: 'P', product: 'powerpoint' },
	visio: { tile: 'V', product: 'visio' },
	openteams: { tile: 'T', product: 'teams' },
};

export const tile = (area) => {
	const meta = AREAS[area] ?? AREAS.core;
	const inner = meta.tile ? escapeHtml(meta.tile) : '<i></i><i></i><i></i><i></i>';
	return `<span class="tile tile-${meta.product}${meta.tile ? '' : ' tile-symbol'}" aria-hidden="true">${inner}</span>`;
};

const isOnly = (block, pattern) => block?.type === 'paragraph' && pattern.test(block.text.trim());
const BOLD_LINE = /^\*\*([^*\n]+)\*\*$/;
const ITALIC_LINE = /^_([^\n]+)_$|^\*([^*\n]+)\*$/;
const IMAGE_LINE = /^!\[[^\]]*\]\([^)]*\)$/;
const IMAGE_PARTS = /^!\[([^\]]*)\]\(\s*([^\s)]+)(?:\s+"[^"]*")?\s*\)$/;
const STAT = /^\*\*([^*\n]+?):\*\*\s*(\S[^\n]*)$/;
/** A feature's limits (release-edition.mjs): a paragraph, never the stat line of a picture. */
const NOT_YET = /^\*\*Not yet:\*\*/;
const isStats = (block) =>
	block.type === 'paragraph' &&
	!NOT_YET.test(block.text.trim()) &&
	block.text
		.trim()
		.split(/\s+·\s+/)
		.every((part) => STAT.test(part));
const italicText = (block) => {
	const m = ITALIC_LINE.exec(block.text.trim());
	return m[1] ?? m[2];
};

/** Join the label, code and caption lines (and image and caption) into single blocks. */
function combine(blocks) {
	const out = [];
	for (let i = 0; i < blocks.length; i++) {
		const block = blocks[i];
		if (isOnly(block, /^(Before|After):$/) && blocks[i + 1]?.type === 'code') {
			out.push({ ...blocks[++i], change: block.text.trim().slice(0, -1) });
		} else if (isOnly(block, BOLD_LINE) && blocks[i + 1]?.type === 'code') {
			const code = { ...blocks[i + 1], file: BOLD_LINE.exec(block.text.trim())[1] };
			i += 1;
			if (isOnly(blocks[i + 1], ITALIC_LINE)) code.caption = italicText(blocks[++i]);
			out.push(code);
		} else if (block.type === 'code' && isOnly(blocks[i + 1], ITALIC_LINE))
			out.push({ ...block, caption: italicText(blocks[++i]) });
		else if (isOnly(block, IMAGE_LINE)) {
			const figure = { type: 'figure', image: block.text.trim() };
			if (isOnly(blocks[i + 1], ITALIC_LINE)) figure.caption = italicText(blocks[++i]);
			if (blocks[i + 1] && isStats(blocks[i + 1])) figure.stats = blocks[++i].text.trim();
			out.push(figure);
		} else if (block.type === 'quote') out.push({ ...block, blocks: combine(block.blocks) });
		else out.push(block);
	}
	return out;
}

/** The page for one post. `link(url, kind)` resolves its URLs; returns the page's HTML. */
export function renderReleasePage(
	source,
	template,
	{
		link = (url) => url,
		media = (url) => ({ src: link(url, 'img') }),
		showcase,
		edition = editionOf(showcase),
		editionDir,
		appearance = showcase?.appearance,
	} = {},
) {
	// A post follows the device's colour scheme until the reader chooses; one designed dark
	// first (`"appearance": "dark"`) opens dark instead. A stored choice wins either way.
	const darkFirst = appearance === 'dark';
	const blocks = combine(parseBlocks(source));
	const take = (index) => blocks.splice(index, 1)[0];
	const titleAt = blocks.findIndex((b) => b.type === 'heading' && b.level === 1);
	const title = titleAt >= 0 ? plainText(take(titleAt).text) : 'Release notes';
	const start = Math.max(titleAt, 0);
	const meta = isOnly(blocks[start], /\*\*Published:\*\*/) ? take(start) : null;
	const hero = blocks[start]?.type === 'figure' ? take(start) : null;
	const heroIntro = showcase?.sections && blocks[start]?.type === 'paragraph' ? take(start) : null;

	const nextSlug = slugger();
	const ctx = { link, slug: nextSlug };
	const show = (type, notes = []) => renderShowcase(showcase, type, { media, tile, notes });
	const toc = [];
	let openDetails = 0;
	let area = null;
	let pillarsDone = false;
	let inLede = true;
	const parts = ['<div class="post-lede">'];
	// An edition wraps each feature (a `###` and what follows it): `out` is where blocks go.
	let out = parts;
	let feature = null;
	const closeFeature = () => {
		if (!feature) return;
		parts.push(renderFeature(feature));
		feature = null;
		out = parts;
	};
	const CHAPTER_LINKS = (index) => `<!--chapter-links:${index}-->`;

	const renderCode = (block) => {
		const lang = LANGUAGE_NAMES[block.lang] ?? (block.lang ? block.lang.toUpperCase() : '');
		const name = block.change ?? block.file ?? lang ?? '';
		const change = block.change ? ` data-migration="${block.change.toLowerCase()}"` : '';
		return [
			`<figure class="code${block.change ? '' : ' reveal'}"${change}>`,
			'<div class="code-bar">',
			`<span class="code-file">${escapeHtml(name || 'Code')}</span>`,
			(block.file || block.change) && lang
				? `<span class="code-lang">${escapeHtml(lang)}</span>`
				: '',
			'<button class="code-copy" type="button" hidden>Copy</button>',
			'</div>',
			`<pre tabindex="0"><code${block.lang ? ` class="language-${escapeHtml(block.lang)}"` : ''}>${highlight(block.text, block.lang)}</code></pre>`,
			block.caption ? `<figcaption>${renderInline(block.caption, ctx)}</figcaption>` : '',
			'</figure>',
		].join('');
	};
	const renderFigure = (block, hero = false) => {
		const [, alt = '', url = ''] = IMAGE_PARTS.exec(block.image) ?? [];
		const info = media(url);
		const size = info.width && info.height ? ` width="${info.width}" height="${info.height}"` : '';
		const load = hero ? ' fetchpriority="high"' : ' loading="lazy" decoding="async"';
		const animated = info.animation ? ` data-animation="${escapeHtml(info.animation)}"` : '';
		const img = (extra = '') =>
			`<img${extra} src="${escapeHtml(info.src)}" alt="${escapeHtml(alt)}"${animated}${size}${load} />`;
		let picture = img();
		if (info.video) {
			const poster = info.poster ? ` poster="${escapeHtml(info.poster)}"` : '';
			const videoClass = info.animated ? ' class="media-still"' : '';
			const video = `<video${videoClass} controls muted playsinline preload="none"${size}${poster} aria-label="${escapeHtml(alt)}"><source src="${escapeHtml(info.video)}" type="video/mp4" /></video>`;
			picture = info.animated
				? `<div class="media-frame">${img(' class="media-motion"')}${video}</div>`
				: video;
		}
		const stats = block.stats
			? `<dl class="post-stats">${block.stats
					.split(/\s+·\s+/)
					.map((part) => {
						const [, label, value] = STAT.exec(part);
						return `<div data-stat="${escapeHtml(slug(label))}"><dt>${renderInline(label, ctx)}</dt><dd>${renderInline(value, ctx)}</dd></div>`;
					})
					.join('')}</dl>`
			: '';
		const caption = block.caption
			? `<figcaption>${renderInline(block.caption, ctx)}</figcaption>`
			: '';
		const className = hero
			? 'post-hero-figure'
			: `post-figure reveal${info.video ? ' post-media' : ''}`;
		return `<figure class="${className}">${picture}${caption}${stats}</figure>`;
	};
	const renderHtml = (html) => {
		let out = html.replace(/<a\s+id="changes-([\w-]+)"\s*><\/a>/g, (match, id) => {
			area = id;
			return `<a id="changes-${escapeHtml(id)}" class="changes-anchor"></a>`;
		});
		out = out.replace(/<details(\s[^>]*)?>/g, (match, attrs = '') => {
			openDetails += 1;
			if (!area) return match;
			return `<details class="changes reveal" data-area="${escapeHtml(area)}"${attrs}>`;
		});
		out = out.replace(/<summary>([\s\S]*?)<\/summary>/g, (match, label) =>
			area
				? `<summary>${tile(area)}<span class="changes-name">${label}</span><span class="changes-count" data-count></span><span class="changes-chevron" aria-hidden="true"></span></summary><div class="changes-body">`
				: match,
		);
		out = out.replace(/<\/details>/g, () => {
			openDetails = Math.max(0, openDetails - 1);
			const close = area ? '</div></details>' : '</details>';
			area = null;
			return close;
		});
		return out;
	};

	const renderChapterBlocks = (items) =>
		items
			.map((block) => {
				if (block.type === 'code') return renderCode(block);
				if (block.type === 'figure') return renderFigure(block);
				return renderBlocks([block], { ...ctx, code: renderCode, html: renderHtml });
			})
			.join('\n');

	for (let i = 0; i < blocks.length; i++) {
		const block = blocks[i];
		if (block.type === 'html') {
			const slot = /^<div data-showcase="([a-z-]+)"><\/div>$/.exec(block.html);
			if (slot) {
				const type = slot[1];
				if (showcase?.[type]) {
					closeFeature();
					let end = i + 1;
					while (
						end < blocks.length &&
						!(blocks[end].type === 'heading' && blocks[end].level === 2)
					)
						end++;
					const { notes, common } = chapterNotes(blocks.slice(i + 1, end), showcase[type]);
					// A recording the scene already shows stays in the Markdown for GitHub only.
					const shown = new Set(showcase[type].map((item) => item.image).filter(Boolean));
					const own = (block) =>
						!(
							block.type === 'figure' &&
							shown.has(basename(IMAGE_PARTS.exec(block.image)?.[2] ?? ''))
						);
					parts.push(
						show(
							type,
							notes.map((note) => renderChapterBlocks(note.filter(own))),
						),
					);
					if (common.length)
						parts.push(`<div class="chapter-common">${renderChapterBlocks(common)}</div>`);
					i = end - 1;
				}
				continue;
			}
		}
		if (block.type === 'heading' && block.level === 2 && openDetails === 0) {
			closeFeature();
			if (inLede) parts.push('</div>');
			else parts.push('</section>');
			inLede = false;
			const plain = plainText(block.text);
			const id = nextSlug(plain);
			toc.push({ id, plain, text: renderInline(block.text, ctx), children: [] });
			const title = heading(2, id, renderInline(block.text, ctx));
			if (edition) {
				const product = chapterProduct(showcase, id);
				parts.push(
					`<section class="post-section"${product ? ` data-product="${product}"` : ''} aria-labelledby="${escapeHtml(id)}">`,
					product
						? `<header class="chapter-head">${chapterTile(product, tile)}${title}${CHAPTER_LINKS(toc.length - 1)}</header>`
						: `<header class="chapter-head">${title}</header>`,
				);
			} else
				parts.push(
					`<section class="post-section reveal" aria-labelledby="${escapeHtml(id)}">`,
					title,
				);
			continue;
		}
		if (block.type === 'heading') {
			const plain = plainText(block.text);
			const id = nextSlug(plain);
			const isFeature = block.level === 3 && openDetails === 0 && toc.length > 0;
			if (isFeature)
				toc[toc.length - 1].children.push({ id, plain, text: renderInline(block.text, ctx) });
			const title = heading(block.level, id, renderInline(block.text, ctx));
			if (edition && isFeature) {
				closeFeature();
				feature = { head: title, media: [], text: [], aside: [] };
				out = feature.text;
			} else out.push(title);
		} else if (feature && block.type === 'figure' && !feature.text.length && !feature.media.length)
			feature.media.push(renderFigure(block));
		else if (feature && isOnly(block, /^(?:\*\*Not yet:\*\*|Learn more\b)/))
			feature.aside.push(
				isOnly(block, /^Learn more\b/)
					? `<p class="learn-more">${renderInline(block.text, ctx)}</p>`
					: `<p class="not-yet">${renderInline(block.text, ctx)}</p>`,
			);
		else if (block.type === 'code') {
			const after = blocks[i + 1];
			if (block.change === 'Before' && after?.type === 'code' && after.change === 'After') {
				out.push(`<div class="migration-pair">${renderCode(block)}${renderCode(after)}</div>`);
				i++;
			} else out.push(renderCode(block));
		} else if (block.type === 'figure') out.push(renderFigure(block));
		else if (block.type === 'html') out.push(renderHtml(block.html));
		else if (
			inLede &&
			!pillarsDone &&
			block.type === 'list' &&
			block.items.every((item) => item.indent === 0 && item.text.startsWith('**'))
		) {
			pillarsDone = true;
			out.push(
				`<ul class="pillars reveal">${block.items.map((item) => `<li>${renderInline(item.text.replace(/\n/g, ' '), ctx)}</li>`).join('')}</ul>`,
			);
		} else if (isOnly(block, /^Learn more\b/))
			out.push(`<p class="learn-more">${renderInline(block.text, ctx)}</p>`);
		else
			out.push(
				renderBlocks([block], {
					...ctx,
					code: renderCode,
					html: renderHtml,
				}),
			);
	}
	closeFeature();
	parts.push(inLede ? '</div>' : '</section>');
	const body = countChanges(parts.join('\n')).replace(
		/<!--chapter-links:(\d+)-->/g,
		(match, index) => renderChapterLinks(toc[Number(index)]),
	);

	const metaHtml = meta
		? renderInline(meta.text, ctx)
				.split(/\s+\|\s+/)
				.map((item) => `<span>${datetime(item)}</span>`)
				.join('')
		: '';
	const lede = heroIntro ?? blocks.find((b) => b.type === 'paragraph' && !b.text.startsWith('**'));
	const description = lede ? plainText(lede.text).replace(/\s+/g, ' ') : title;
	const tocHtml = toc
		.map(
			(entry) =>
				`<li><a href="#${escapeHtml(entry.id)}">${stripLinks(entry.text)}</a>${
					entry.children.length
						? `<ul>${entry.children.map((child) => `<li><a href="#${escapeHtml(child.id)}">${stripLinks(child.text)}</a></li>`).join('')}</ul>`
						: ''
				}</li>`,
		)
		.join('');
	const fill = {
		title: escapeHtml(title),
		titleHtml: escapeHtml(title).replace(
			/(\s)(\d+(?:\.\d+)*)$/,
			'$1<span class="post-version">$2</span>',
		),
		description: escapeHtml(
			description.length > 200 ? `${description.slice(0, 197)}...` : description,
		),
		meta: metaHtml,
		// An edition lays the page out itself; `post-indexed` says its hero holds the contents.
		eventClass: edition
			? showcase?.sections
				? ' post-edition post-indexed'
				: ' post-edition'
			: showcase?.sections
				? ' release-event'
				: '',
		appearance: darkFirst ? 'dark' : 'system',
		appearanceTitle: darkFirst
			? 'Dark mode: on for this post until you choose'
			: 'Dark mode: follows your device until you choose',
		edition: edition ? ` data-edition="${escapeHtml(edition)}"` : '',
		editionAssets: editionAssets(edition, editionDir),
		heroIntro: heroIntro ? `<p class="release-intro">${renderInline(heroIntro.text, ctx)}</p>` : '',
		releaseIndex: edition
			? renderEditionIndex(showcase, toc, tile)
			: renderReleaseIndex(showcase?.sections, showcase?.version),
		hero: show('shell') || (hero ? renderFigure(hero, true) : ''),
		toc: tocHtml,
		body,
	};
	// The two edition slots take the white space before them, so a post without an edition is
	// byte for byte the page it was before editions (`<body {{edition}}>` renders as `<body>`).
	return template
		.replace(/\s*\{\{(edition|editionAssets)\}\}/g, (match, key) => fill[key])
		.replace(/\{\{(\w+)\}\}/g, (match, key) => (key in fill ? fill[key] : match));
}

function chapterNotes(blocks, items) {
	const groups = new Map();
	let current = [];
	groups.set('', current);
	for (const block of blocks) {
		if (block.type === 'heading' && block.level === 3) {
			current = [];
			groups.set(plainText(block.text), current);
		}
		current.push(block);
	}
	const notes = items.map(({ notes = [] }) =>
		notes.flatMap((name) => {
			const group = groups.get(name);
			if (!group) throw Error(`No chapter heading named ${name}`);
			groups.delete(name);
			return group;
		}),
	);
	return { notes, common: [...groups.values()].flat() };
}

const heading = (level, id, inner) => `<h${level} id="${escapeHtml(id)}">${inner}</h${level}>`;
const stripLinks = (html) => html.replace(/<\/?a\b[^>]*>/g, '');

/** `<strong>Published:</strong> November 21, 2026` with the date in a `<time>` element. */
function datetime(item) {
	return item.replace(/^(<strong>Published:<\/strong>\s*)(.+)$/, (match, label, date) => {
		const parsed = new Date(`${date} 12:00 UTC`);
		return Number.isNaN(parsed.getTime())
			? match
			: `${label}<time datetime="${parsed.toISOString().slice(0, 10)}">${date}</time>`;
	});
}

/** Fill each change list's count with the bullets it holds. */
function countChanges(html) {
	return html.replace(
		/(<details class="changes[\s\S]*?<span class="changes-count" data-count>)(<\/span>)([\s\S]*?<\/details>)/g,
		(match, head, close, rest) => {
			const count = (rest.match(/<li>/g) ?? []).length;
			return count
				? `${head}${count === 1 ? '1 change' : `${count} changes`}${close}${rest}`
				: match;
		},
	);
}

/** Every release post and its optional migration guide, with shared assets. Returns post names. */
export function writeReleasePages(out, { root = ROOT } = {}) {
	const dir = join(root, POSTS_DIR);
	const template = readFileSync(join(root, TEMPLATE), 'utf8');
	const shared = join(root, dirname(TEMPLATE));
	mkdirSync(join(out, 'releases'), { recursive: true });
	for (const file of readdirSync(shared))
		if (file !== basename(TEMPLATE)) copyFileSync(join(shared, file), join(out, 'releases', file));
	// The template is source, not a page (build-pages.mjs copies the whole of site/).
	rmSync(join(out, 'releases', basename(TEMPLATE)), { force: true });
	const names = [];
	for (const file of existsSync(dir) ? readdirSync(dir) : []) {
		if (extname(file) !== '.md') continue;
		const name = basename(file, '.md');
		const assets = new Map();
		const target = join(out, 'releases', name);
		mkdirSync(join(target, 'assets'), { recursive: true });
		const guide = `docs/migration-${name}.md`;
		const hasGuide = existsSync(join(root, guide));
		const config = join(dir, name, 'showcase.json');
		const showcase = existsSync(config) ? JSON.parse(readFileSync(config, 'utf8')) : undefined;
		const render = (source, relatedPage, show) => {
			const path = join(root, source);
			return renderReleasePage(readFileSync(path, 'utf8'), template, {
				...mediaResolver({ root, postsDir: dirname(path), name, assets, relatedPage }),
				showcase: show,
				// The migration guide wears its post's edition.
				edition: editionOf(showcase),
				editionDir: shared,
				appearance: showcase?.appearance,
			});
		};
		const post = `${POSTS_DIR}/${file}`;
		writeFileSync(
			join(target, 'index.html'),
			render(post, hasGuide ? { source: guide, href: 'migration.html' } : undefined, showcase),
		);
		if (hasGuide)
			writeFileSync(join(target, 'migration.html'), render(guide, { source: post, href: './' }));
		for (const [to, from] of assets) {
			if (!existsSync(from)) throw new Error(`${file}: the picture ${from} does not exist`);
			copyFileSync(from, join(target, to));
		}
		names.push(name);
	}
	return names;
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
	const out = resolve(ROOT, process.argv[2] ?? 'suite-dist');
	const names = writeReleasePages(out);
	console.log(`Release pages: ${names.map((name) => `releases/${name}/`).join(', ') || 'none'}`);
}
