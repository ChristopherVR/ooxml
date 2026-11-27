import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

import { TEMPLATE, renderReleasePage, writeReleasePages } from './build-release-notes.mjs';
import { highlight } from './release-highlight.mjs';
import { imageSize } from './release-media.mjs';
import {
	escapeHtml,
	plainText,
	renderInline,
	renderMarkdown,
	slug,
	slugger,
} from './release-markdown.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const TEMPLATE_HTML =
	'<title>{{title}}</title><h1>{{titleHtml}}</h1><p>{{meta}}</p>{{hero}}<ul>{{toc}}</ul><main>{{body}}</main><meta content="{{description}}">';

test('slugs follow GitHub: lower case, punctuation dropped, spaces to hyphens, repeats numbered', () => {
	assert.equal(slug('Platform: ooxml-core'), 'platform-ooxml-core');
	assert.equal(slug('Shared UI and MCP'), 'shared-ui-and-mcp');
	assert.equal(slug('Every change, by product'), 'every-change-by-product');
	assert.equal(slug('Excel calculation (ooxml-core/xlsx)'), 'excel-calculation-ooxml-corexlsx');
	const next = slugger();
	assert.deepEqual([next('Fixed'), next('Fixed'), next('Fixed')], ['fixed', 'fixed-1', 'fixed-2']);
	assert.equal(
		plainText('Use `CalcEngine` with **care** and [links](x)'),
		'Use CalcEngine with care and links',
	);
});

test('inline markup renders and every other character is escaped', () => {
	assert.equal(
		renderInline('**Bold:** _it_ and *it* with `a<b>` and [x & y](https://e.test/?a=1&b=2)'),
		'<strong>Bold:</strong> <em>it</em> and <em>it</em> with <code>a&lt;b&gt;</code> and <a href="https://e.test/?a=1&amp;b=2" rel="noopener">x &amp; y</a>',
	);
	assert.equal(renderInline('<script>alert(1)</script>'), '&lt;script&gt;alert(1)&lt;/script&gt;');
	assert.equal(renderInline('snake_case_name stays'), 'snake_case_name stays');
	assert.equal(renderInline('\\*not em\\*'), '*not em*');
	assert.equal(
		renderInline('[\\] and `code]`](guide.md)'),
		'<a href="guide.md">] and <code>code]</code></a>',
	);
	assert.equal(renderInline('[bad](javascript:alert(1))'), '<a href="#">bad</a>)');
	assert.match(renderInline('![a "b"](pic.png)'), /<img src="pic.png" alt="a &quot;b&quot;"/);
	assert.equal(escapeHtml(`'"&`), '&#39;&quot;&amp;');
});

test('blocks: headings, nested lists, tables, code, quotes, alerts, comments and raw HTML', () => {
	const html = renderMarkdown(
		[
			'## Word',
			'',
			'- one',
			'  - nested **bold**',
			'- two',
			'',
			'| A | B |',
			'| - | :-: |',
			'| `x` | y |',
			'',
			'```ts',
			'const a = "<b>";',
			'```',
			'',
			'> [!WARNING]',
			'> Mind the gap.',
			'',
			'> Plain quote.',
			'',
			'<!-- hidden',
			'comment -->',
			'<details>',
			'<summary>Raw</summary>',
			'',
			'Inside.',
			'',
			'</details>',
		].join('\n'),
	);
	assert.match(html, /<h2 id="word">Word<\/h2>/);
	assert.match(
		html,
		/<ul><li>one<ul><li>nested <strong>bold<\/strong><\/li><\/ul><\/li><li>two<\/li><\/ul>/,
	);
	assert.match(html, /<th>A<\/th><th style="text-align:center">B<\/th>/);
	assert.match(html, /<td><code>x<\/code><\/td>/);
	// A wide table scrolls in its wrapper, which the keyboard can reach.
	assert.match(html, /<div class="table-wrap" tabindex="0"><table>/);
	assert.match(html, /<code class="language-ts">const a = &quot;&lt;b&gt;&quot;;<\/code>/);
	assert.match(
		html,
		/<div class="callout callout-warning" role="note"><p class="callout-title">Warning<\/p><p>Mind the gap.<\/p><\/div>/,
	);
	assert.match(html, /<blockquote><p>Plain quote.<\/p><\/blockquote>/);
	assert.doesNotMatch(html, /hidden/);
	assert.match(html, /<details>\n<summary>Raw<\/summary>\n<p>Inside.<\/p>\n<\/details>/);
});

test('the highlighter marks tokens and escapes the rest', () => {
	assert.equal(
		highlight('const x = "<a>"; // hi', 'ts'),
		'<span class="tok-keyword">const</span> x = <span class="tok-string">&quot;&lt;a&gt;&quot;</span>; <span class="tok-comment">// hi</span>',
	);
	assert.equal(
		highlight('bun add ooxml-core@^2 --exact', 'sh'),
		'<span class="tok-command">bun</span> add ooxml-core@^2 <span class="tok-attr">--exact</span>',
	);
	assert.equal(highlight('<b>', 'unknown'), '&lt;b&gt;');
});

const POST = `# OOXML 2.0

**Published:** November 21, 2026 | **Author:** [Someone](https://github.com/someone)

![Hero](hero.png)

The lede paragraph.

- **Safe:** one.
- **Fast:** two.

\`\`\`sh
bun add ooxml-core@^2
\`\`\`

> [!NOTE]
> Read the [guide](../migration-2.0.md).

## Faster files

### Word saves

Text.

**save.ts**

\`\`\`ts
save();
\`\`\`

_What it does._

![A recording](2.0/assets/clip.gif)

_Before and after._

**Before:** 4.8 s · **After:** 0.6 s

Learn more about [saving](../performance-budgets.md).

## Every change, by product

<a id="changes-word"></a>
<details>
<summary>Word</summary>

### Fixed

- One.
- Two.

#### Hidden heading

</details>
`;

test('the page takes the title, metadata, hero, lede pillars, sections and change lists', () => {
	const media = (url) =>
		url.endsWith('.gif')
			? {
					src: 'assets/clip.gif',
					width: 960,
					height: 540,
					video: 'assets/clip.mp4',
					animated: true,
				}
			: { src: `assets/${url}`, width: 1200, height: 800 };
	const html = renderReleasePage(POST, TEMPLATE_HTML, { media });
	assert.match(html, /<title>OOXML 2.0<\/title>/);
	assert.match(html, /<h1>OOXML <span class="post-version">2.0<\/span><\/h1>/);
	assert.match(
		html,
		/<span><strong>Published:<\/strong> <time datetime="2026-11-21">November 21, 2026<\/time><\/span><span><strong>Author:<\/strong>/,
	);
	assert.match(
		html,
		/<figure class="post-hero-figure"><img src="assets\/hero.png" alt="Hero" width="1200" height="800" fetchpriority="high" \/>/,
	);
	assert.match(html, /content="The lede paragraph."/);
	assert.match(html, /<ul class="pillars reveal"><li><strong>Safe:<\/strong> one.<\/li>/);
	assert.match(html, /<span class="code-file">Shell<\/span>/);
	assert.match(html, /callout-note/);
	assert.match(html, /<section class="post-section reveal" aria-labelledby="faster-files">/);
	assert.match(
		html,
		/<span class="code-file">save.ts<\/span><span class="code-lang">TypeScript<\/span>/,
	);
	assert.match(html, /<figcaption>What it does.<\/figcaption><\/figure>/);
	assert.match(
		html,
		/class="media-motion" src="assets\/clip.gif" alt="A recording" width="960" height="540" loading="lazy"/,
	);
	assert.match(
		html,
		/<video class="media-still" controls muted playsinline preload="none" width="960" height="540"/,
	);
	assert.match(
		html,
		/<dl class="post-stats"><div data-stat="before"><dt>Before<\/dt><dd>4.8 s<\/dd><\/div><div data-stat="after">/,
	);
	assert.match(html, /<p class="learn-more">Learn more about/);
	assert.match(html, /<a id="changes-word" class="changes-anchor"><\/a>/);
	assert.match(html, /<details class="changes reveal" data-area="word">/);
	assert.match(
		html,
		/<span class="tile tile-word" aria-hidden="true">W<\/span><span class="changes-name">Word<\/span><span class="changes-count" data-count>2 changes<\/span>/,
	);
	// The contents list the sections and their features, not the headings inside a change list.
	assert.match(
		html,
		/<li><a href="#faster-files">Faster files<\/a><ul><li><a href="#word-saves">Word saves<\/a><\/li><\/ul><\/li>/,
	);
	assert.match(html, /<li><a href="#every-change-by-product">Every change, by product<\/a><\/li>/);
	assert.doesNotMatch(html, /href="#fixed"/);
	assert.match(html, /<h4 id="hidden-heading">/);
});

test('picture sizes are read from PNG, GIF and SVG headers', () => {
	const png = Buffer.alloc(24);
	png.writeUInt32BE(0x89504e47, 0);
	png.writeUInt32BE(640, 16);
	png.writeUInt32BE(360, 20);
	assert.deepEqual(imageSize(png), { width: 640, height: 360 });
	const gif = Buffer.from('GIF89a\xc0\x03\x1c\x02', 'latin1');
	assert.deepEqual(imageSize(gif), { width: 960, height: 540 });
	assert.deepEqual(imageSize(Buffer.from('<svg viewBox="0 0 960 400">')), {
		width: 960,
		height: 400,
	});
	assert.equal(imageSize(Buffer.from('nothing')), null);
});

test('writeReleasePages renders every post and copies its media next to the page', () => {
	const root = mkdtempSync(join(tmpdir(), 'release-notes-'));
	const out = join(root, 'out');
	try {
		mkdirSync(join(root, 'docs/releases/2.0/assets'), { recursive: true });
		mkdirSync(join(root, dirname(TEMPLATE)), { recursive: true });
		writeFileSync(join(root, TEMPLATE), TEMPLATE_HTML);
		writeFileSync(join(root, 'site/releases/releases.css'), 'body{}');
		writeFileSync(
			join(root, 'docs/releases/2.0/assets/clip.gif'),
			Buffer.from('GIF89a\xc0\x03\x1c\x02', 'latin1'),
		);
		writeFileSync(join(root, 'docs/releases/2.0/assets/clip.mp4'), 'mp4');
		writeFileSync(join(root, 'docs/releases/2.0/assets/clip.poster.png'), 'png');
		writeFileSync(
			join(root, 'docs/releases/2.0.md'),
			'# OOXML 2.0\n\nLede.\n\n## Fast\n\n![Clip](assets/clip.gif)\n\nSee [guide](../migration-2.0.md#top).\n',
		);
		assert.deepEqual(writeReleasePages(out, { root }), ['2.0']);
		const html = readFileSync(join(out, 'releases/2.0/index.html'), 'utf8');
		assert.match(html, /src="assets\/clip.gif" alt="Clip" width="960" height="540"/);
		assert.match(html, /poster="assets\/clip.poster.png"/);
		assert.match(
			html,
			/href="https:\/\/github.com\/ChristopherVR\/ooxml\/blob\/main\/docs\/migration-2.0.md#top"/,
		);
		assert.equal(readFileSync(join(out, 'releases/2.0/assets/clip.mp4'), 'utf8'), 'mp4');
		assert.equal(readFileSync(join(out, 'releases/releases.css'), 'utf8'), 'body{}');
		assert.throws(() => readFileSync(join(out, 'releases/template.html')));
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
});

for (const post of ['2.0', '2.5'])
	test(`the ${post} post in the repository renders with its stable change anchors`, () => {
		const html = renderReleasePage(
			readFileSync(join(ROOT, `docs/releases/${post}.md`), 'utf8'),
			readFileSync(join(ROOT, TEMPLATE), 'utf8'),
		);
		for (const area of ['core', 'ui-mcp', 'word', 'excel', 'powerpoint', 'visio', 'openteams'])
			assert.match(
				html,
				new RegExp(
					`<a id="changes-${area}" class="changes-anchor"></a>\\s*<details class="changes reveal" data-area="${area}">`,
				),
			);
		assert.doesNotMatch(html, /\{\{\w+\}\}/);
		assert.doesNotMatch(html, /\u2014/u);
	});

test('the 2.5 post shows its eight recordings once, from posters, with its own version', () => {
	const source = readFileSync(join(ROOT, 'docs/releases/2.5.md'), 'utf8');
	const showcase = JSON.parse(readFileSync(join(ROOT, 'docs/releases/2.5/showcase.json'), 'utf8'));
	const assets = join(ROOT, 'docs/releases/2.5/assets');
	const media = (url) => {
		const file = url.split('/').pop();
		const stem = file.replace(/\.gif$/u, '');
		return file.endsWith('.gif')
			? { src: `assets/${stem}.poster.webp`, animation: `assets/${file}` }
			: { src: `assets/${file}` };
	};
	const template = readFileSync(join(ROOT, TEMPLATE), 'utf8');
	const html = renderReleasePage(source, template, { showcase, media });
	assert.match(html, /<html lang="en" data-appearance-default="system">/);
	assert.match(html, /aria-label="2\.5 highlights"/);
	assert.doesNotMatch(html, /2\.0 highlights|Before 2\.0/);
	for (const { id } of showcase.sections) assert.ok(html.includes(`<h2 id="${id}">`), id);
	const items = showcase['new-features'];
	assert.equal(items.length, 8);
	for (const item of items) {
		const gif = readFileSync(join(assets, item.image));
		assert.equal(gif.toString('ascii', 0, 3), 'GIF', item.image);
		assert.ok(gif.length <= 1024 * 1024, `${item.image} is ${gif.length} bytes`);
		assert.deepEqual(imageSize(gif), { width: 1180, height: 802 }, item.image);
		readFileSync(join(assets, item.image.replace(/\.gif$/u, '.poster.webp')));
		// The scene draws the recording; the Markdown keeps it for GitHub, so it must not repeat.
		assert.equal(html.split(`data-animation="assets/${item.image}"`).length - 1, 1, item.image);
		const alt = new RegExp(
			`!\\[([^\\]]+)\\]\\(2\\.5/assets/${item.image.replace('.', '\\.')}\\)`,
			'u',
		);
		assert.match(source, alt, `${item.image} has alt text in the Markdown`);
	}
	assert.match(html, /class="tile tile-visio"/);
	assert.doesNotMatch(html, /<video\b/);
});

test('still images offer controlled playback and the release guide links back to the post', () => {
	const root = mkdtempSync(join(tmpdir(), 'release-guide-'));
	const out = join(root, 'out');
	try {
		mkdirSync(join(root, 'docs/releases/2.0/assets'), { recursive: true });
		mkdirSync(join(root, dirname(TEMPLATE)), { recursive: true });
		writeFileSync(join(root, TEMPLATE), TEMPLATE_HTML);
		writeFileSync(join(root, 'docs/releases/2.0/assets/demo.svg'), '<svg viewBox="0 0 960 540">');
		writeFileSync(join(root, 'docs/releases/2.0/assets/demo.mp4'), 'mp4');
		writeFileSync(
			join(root, 'docs/releases/2.0.md'),
			'# OOXML 2.0\n\nRead the [guide](../migration-2.0.md#upgrade).\n\n## Demo\n\n![Exporting a slide](2.0/assets/demo.svg)\n\n_[Video](2.0/assets/demo.mp4)._\n',
		);
		writeFileSync(
			join(root, 'docs/migration-2.0.md'),
			'# Migration\n\nRead the [post](releases/2.0.md#demo).\n\n## Upgrade\n',
		);
		assert.deepEqual(writeReleasePages(out, { root }), ['2.0']);
		const blog = readFileSync(join(out, 'releases/2.0/index.html'), 'utf8');
		const guide = readFileSync(join(out, 'releases/2.0/migration.html'), 'utf8');
		assert.match(blog, /href="migration.html#upgrade"/);
		assert.match(
			blog,
			/<video controls muted playsinline preload="none" width="960" height="540" poster="assets\/demo.svg" aria-label="Exporting a slide">/,
		);
		assert.match(blog, /<figcaption><a href="assets\/demo.mp4">Video<\/a>\.<\/figcaption>/);
		assert.doesNotMatch(blog, /autoplay|media-motion|media-still/);
		assert.match(guide, /href="\.\/#demo"/);
		assert.match(guide, /<h2 id="upgrade">/);
		assert.equal(readFileSync(join(out, 'releases/2.0/assets/demo.mp4'), 'utf8'), 'mp4');
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
});

test('the 2.0 post leads with its release and renders the five interactive chapters', () => {
	const source = readFileSync(join(ROOT, 'docs/releases/2.0.md'), 'utf8');
	const showcase = JSON.parse(readFileSync(join(ROOT, 'docs/releases/2.0/showcase.json'), 'utf8'));
	const html = renderReleasePage(source, readFileSync(join(ROOT, TEMPLATE), 'utf8'), { showcase });
	assert.match(html, /aria-label="Install OOXML 2.0"/);
	// 2.0 follows the device until the reader chooses.
	assert.match(html, /<html lang="en" data-appearance-default="system">/);
	assert.match(html, /data-copy-command="npm install pptx-react-viewer@5 react"/);
	assert.match(html, /class="release-intro"/);
	assert.doesNotMatch(html, /OOXML is an open-source Office stack/);
	for (const { id } of showcase.sections) {
		assert.ok(html.includes(`href="#${id}"`));
		assert.ok(html.includes(`data-showcase="${id}"`));
	}
	assert.doesNotMatch(html, /class="showcase-details"/);
	assert.match(html, /data-note="1" hidden>[^]*id="export-and-downloads-on-safari"/);
	assert.match(html, /aria-label="Performance comparisons"/);
	const forPerformance = html.slice(
		html.indexOf('data-showcase="performance"'),
		html.indexOf('<div class="showcase-notes"', html.indexOf('data-showcase="performance"')),
	);
	assert.doesNotMatch(forPerformance, /<img/);
	assert.match(forPerformance, /Click Save/);
	assert.match(html, /<h3>Save a one-paragraph edit<\/h3>/);
	assert.match(html, /<dt>Before 2\.0<\/dt><dd style="[^"]+">13\.53 s<\/dd>/);
	assert.match(html, /<dt>OOXML 2\.0<\/dt><dd style="[^"]+">1\.97 s<\/dd>/);
	assert.match(html, /<dt>Before 2\.0<\/dt><dd style="[^"]+">Drawing refused<\/dd>/);
	assert.match(html, /signature-invalid\.webp/);
	assert.match(html, /private-chat-other\.webp/);
	assert.match(html, /limits-refused\.webp/);
	assert.match(html, /A2: the formula was refused because it uses WEBSERVICE/);
	assert.doesNotMatch(html, /powerpoint-webkit|word-save\.mp4/);
	assert.doesNotMatch(html, /<video\b/);
});

test('a GIF with a poster starts as a still without playback buttons', () => {
	const { media } = {
		media: () => ({
			src: 'assets/demo.poster.webp',
			animation: 'assets/demo.gif',
			width: 960,
			height: 540,
		}),
	};
	const html = renderReleasePage(
		'# Demo\n\n## Motion\n\n![Format a chart](demo.gif)',
		TEMPLATE_HTML,
		{ media },
	);
	assert.match(
		html,
		/<img src="assets\/demo\.poster\.webp" alt="Format a chart" data-animation="assets\/demo\.gif"/,
	);
	assert.doesNotMatch(html, /<button/);
	assert.doesNotMatch(html, /<video\b/);
});

test('the migration guide labels and groups before-and-after code examples', () => {
	const source =
		'# Migrate\n\n## Upgrade\n\nBefore:\n\n' +
		'```sh\nnpm install ooxml-core@1\n```\n\nAfter:\n\n```sh\nnpm install ooxml-core@2\n```';
	const html = renderReleasePage(source, TEMPLATE_HTML);
	assert.match(html, /class="migration-pair">/);
	assert.match(html, /data-migration="before"><div[^>]*><span[^>]*>Before<\/span>/);
	assert.match(html, /data-migration="after"><div[^>]*><span[^>]*>After<\/span>/);
});

test('an edition is opt-in: its markup, assets and limits appear only when a post names one', () => {
	const source = [
		'# Demo 3.1',
		'',
		'**Published:** January 2, 2027',
		'',
		'What the release does.',
		'',
		'## Word',
		'',
		'### Word Draw',
		'',
		'![Draws a line](draw.gif)',
		'',
		'Pens.',
		'',
		'**Not yet:** A ruler.',
		'',
		'Learn more about [ink](ink.md).',
		'',
		'### Views',
		'',
		'Read Mode.',
		'',
		'## Security',
		'',
		'Checked.',
	].join('\n');
	const template =
		'<body {{edition}}>{{editionAssets}}<main class="post{{eventClass}}">{{heroIntro}}{{releaseIndex}}{{body}}</main>';
	const plain = renderReleasePage(source, template);
	assert.match(plain, /^<body><main class="post">/);
	assert.match(
		plain,
		/<section class="post-section reveal" aria-labelledby="word">\n<h2 id="word">/,
	);
	assert.doesNotMatch(plain, /data-edition|chapter-head|class="feature|not-yet|ribbon-/);

	const config = {
		version: '3.1',
		edition: 'ribbon',
		sections: [
			{ id: 'word', label: 'Word', product: 'word' },
			{ id: 'security', label: 'Security' },
		],
	};
	const html = renderReleasePage(source, template, { showcase: config });
	assert.match(
		html,
		/^<body data-edition="ribbon">\n\t\t<link rel="stylesheet" href="\.\.\/edition-ribbon\.css" \/><main class="post post-edition post-indexed">/,
	);
	assert.match(html, /<p class="release-intro">What the release does\.<\/p>/);
	assert.match(
		html,
		/<ul class="ribbon-tabs"><li><a href="#word-draw">Draw<\/a><\/li><li><a href="#views">Views<\/a><\/li><\/ul>/,
	);
	assert.match(html, /<ul class="ribbon-tabs"><li><a href="#security">Security<\/a><\/li><\/ul>/);
	assert.match(
		html,
		/<div class="feature"><h3 id="word-draw">Word Draw<\/h3><figure class="post-figure reveal"><img [^>]+><\/figure><div class="feature-body"><div class="feature-text"><p>Pens\.<\/p><\/div><div class="feature-aside"><p class="not-yet"><strong>Not yet:<\/strong> A ruler\.<\/p>\n<p class="learn-more">Learn more about <a href="ink\.md">ink<\/a>\.<\/p><\/div><\/div><\/div>/,
	);
	assert.match(html, /<div class="feature"><h3 id="views">Views<\/h3><div class="feature-body">/);
	// A guide rendered for the post wears the edition without the post's index.
	const guide = renderReleasePage('# Guide\n\n## Upgrade\n\nSteps.', template, {
		edition: 'ribbon',
	});
	assert.match(guide, /<main class="post post-edition">/);
	assert.doesNotMatch(guide, /ribbon-index/);
	// A name that is not a plain file stem is not an edition.
	assert.doesNotMatch(
		renderReleasePage(source, template, { showcase: { edition: '../x' } }),
		/data-edition/,
	);
	// A section the post does not have is a mistake in its configuration.
	assert.throws(
		() =>
			renderReleasePage(source, template, {
				showcase: { edition: 'ribbon', sections: [{ id: 'excel', product: 'excel' }] },
			}),
		/No chapter with the id excel/,
	);
});

test('a post follows the device unless its configuration says it opens dark', () => {
	const template =
		'<html data-appearance-default="{{appearance}}"><button title="{{appearanceTitle}}"></button>{{body}}</html>';
	const source = '# Demo\n\n## One\n\nText.';
	for (const showcase of [undefined, {}, { appearance: 'system' }, { appearance: 'light' }])
		assert.match(
			renderReleasePage(source, template, { showcase }),
			/^<html data-appearance-default="system"><button title="Dark mode: follows your device until you choose">/,
		);
	assert.match(
		renderReleasePage(source, template, { showcase: { appearance: 'dark' } }),
		/^<html data-appearance-default="dark"><button title="Dark mode: on for this post until you choose">/,
	);
	// A migration guide is rendered without its post's showcase and is told the default.
	assert.match(
		renderReleasePage(source, template, { appearance: 'dark' }),
		/^<html data-appearance-default="dark">/,
	);
});
