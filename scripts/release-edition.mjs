/**
 * release-edition.mjs: a release post's own look. A post whose `showcase.json` names an
 * `edition` is rendered by the same builder on the same template, with:
 *
 * - `data-edition="<name>"` on the page's `<body>` and the stylesheet (and script, when there is
 *   one) `site/releases/edition-<name>.css|js` after the shared ones, so every rule of the
 *   edition is scoped to that attribute and no other post can pick it up;
 * - the hero's index drawn as one row per chapter with the chapter's features in it;
 * - each `##` chapter opened by a header that carries its product and its features as tabs;
 * - each `###` feature wrapped, with its limits (`**Not yet:**`) and `Learn more` link set aside.
 *
 * A post without an edition is rendered exactly as before.
 */
import { existsSync } from 'node:fs';
import { join } from 'node:path';

import { escapeHtml } from './release-markdown.mjs';

const NAME = /^[a-z][a-z0-9-]*$/;
const PRODUCTS = new Set(['word', 'excel', 'powerpoint', 'visio', 'teams', 'office']);

/** The edition a post's configuration names, or '' (also for a name that is not a file stem). */
export const editionOf = (config) =>
	typeof config?.edition === 'string' && NAME.test(config.edition) ? config.edition : '';

/** The edition's stylesheet and script tags, for the files that exist in `dir`. */
export function editionAssets(edition, dir) {
	if (!edition) return '';
	const tags = [];
	if (!dir || existsSync(join(dir, `edition-${edition}.css`)))
		tags.push(`<link rel="stylesheet" href="../edition-${edition}.css" />`);
	if (dir && existsSync(join(dir, `edition-${edition}.js`)))
		tags.push(`<script type="module" src="../edition-${edition}.js"></script>`);
	// Each tag brings its own line break, so a post without an edition renders as it always did.
	return tags.map((tag) => `\n\t\t${tag}`).join('');
}

/** The product a chapter belongs to, from the post's `sections` entry with the chapter's id. */
export function chapterProduct(config, id) {
	const product = config?.sections?.find((section) => section.id === id)?.product;
	return PRODUCTS.has(product) ? product : '';
}

/** A product's tile, from the builder's `tile(area)`; the suite's symbol stands for the suite. */
export const chapterTile = (product, tile) =>
	tile(product === 'office' ? 'core' : product === 'teams' ? 'openteams' : product);

/** A feature's tab: its heading without its chapter's name in front ("Word Draw" is "Draw"). */
export function tabLabel(text, chapter) {
	const plain = text.trim();
	return chapter && plain.startsWith(`${chapter} `) && plain.length > chapter.length + 1
		? plain.slice(chapter.length + 1)
		: plain;
}

const tabs = (entry) =>
	entry.children
		.map(
			(child) =>
				`<li><a href="#${escapeHtml(child.id)}">${escapeHtml(tabLabel(child.plain, entry.plain))}</a></li>`,
		)
		.join('');

/** The tabs of one chapter, under its heading: its features, in order. */
export function renderChapterTabs(entry) {
	if (!entry?.children.length) return '';
	return `<nav class="chapter-tabs" aria-label="${escapeHtml(entry.plain)} features"><ul>${tabs(entry)}</ul></nav>`;
}

/**
 * The hero's index: one row per chapter that has a product (its tile, its name and its features
 * as tabs) and a last row with the other chapters the post's `sections` name.
 */
export function renderEditionIndex(config, toc, tile) {
	const sections = config?.sections ?? [];
	const byId = new Map(toc.map((entry) => [entry.id, entry]));
	const rows = [];
	const rest = [];
	for (const section of sections) {
		const entry = byId.get(section.id);
		if (!entry) throw Error(`No chapter with the id ${section.id}`);
		const product = chapterProduct(config, section.id);
		if (!product) {
			rest.push(
				`<li><a href="#${escapeHtml(section.id)}">${escapeHtml(section.label ?? entry.plain)}</a></li>`,
			);
			continue;
		}
		rows.push(
			`<li class="ribbon-row" data-product="${product}"><a class="ribbon-name" href="#${escapeHtml(section.id)}">${chapterTile(product, tile)}<span>${escapeHtml(section.label ?? entry.plain)}</span></a><ul class="ribbon-tabs">${tabs(entry)}</ul></li>`,
		);
	}
	if (rest.length)
		rows.push(
			`<li class="ribbon-row ribbon-rest"><span class="ribbon-name">More</span><ul class="ribbon-tabs">${rest.join('')}</ul></li>`,
		);
	if (!rows.length) return '';
	const version = escapeHtml(config.version ?? '');
	return `<nav class="ribbon-index" aria-label="${version ? `In OOXML ${version}` : 'In this release'}"><ul class="ribbon-rows">${rows.join('')}</ul></nav>`;
}

/** One feature: its heading, its opening recording, then its text beside its limits and links. */
export function renderFeature({ head, media, text, aside }) {
	const side = aside.length ? `<div class="feature-aside">${aside.join('\n')}</div>` : '';
	const body =
		text.length || side
			? `<div class="feature-body"><div class="feature-text">${text.join('\n')}</div>${side}</div>`
			: '';
	return `<div class="feature">${head}${media.join('\n')}${body}</div>`;
}
