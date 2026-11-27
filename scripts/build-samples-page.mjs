#!/usr/bin/env node
/**
 * build-samples-page.mjs: the samples gallery at `<site>/samples/`.
 *
 * The gallery is a static page: `site/samples/template.html` with one card per entry of
 * `samples/manifest.json`, beside the sample files and their thumbnails. It ships no script of
 * its own (the appearance switch, search and reveals are the documentation hub's), so it costs the
 * suite's JavaScript budget nothing. `build-suite.mjs` and `build-pages.mjs` call
 * `writeSamplesPage` after copying `site/`.
 *
 *   node scripts/build-samples-page.mjs [out]      # default: suite-dist
 */
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { SAMPLE_ID } from '../apps/office-suite/src/samples.js';
import { escapeHtml } from './release-markdown.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const TEMPLATE = 'site/samples/template.html';
/** Thumbnails are rendered at this size (`samples/generators/thumbnails.ps1`). */
const THUMBNAIL = { width: 800, height: 450 };

/** The products in gallery order, with the app each file opens in. */
export const PRODUCTS = [
	{ id: 'docx', app: 'Word', letter: 'W', colour: '#185abd', noun: ['document', 'documents'] },
	{ id: 'xlsx', app: 'Excel', letter: 'X', colour: '#107c41', noun: ['workbook', 'workbooks'] },
	{
		id: 'pptx',
		app: 'PowerPoint',
		letter: 'P',
		colour: '#c43e1c',
		noun: ['presentation', 'presentations'],
	},
	{ id: 'visio', app: 'Visio', letter: 'V', colour: '#3955a3', noun: ['drawing', 'drawings'] },
];

const WORDS = [
	'No',
	'One',
	'Two',
	'Three',
	'Four',
	'Five',
	'Six',
	'Seven',
	'Eight',
	'Nine',
	'Ten',
	'Eleven',
	'Twelve',
];
const countWord = (count) => WORDS[count] ?? String(count);

/** `412 KB`, `1.2 MB`: the size a card shows. */
export function formatSize(bytes) {
	if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
	return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * The entries of a parsed manifest, checked for everything the page writes into HTML and the suite
 * reads: a slug id (it becomes an element id and a link), text fields, a list of feature strings,
 * a byte size, a SHA-256 and paths that stay inside samples/.
 */
export function checkManifest(manifest) {
	const fail = (sample, why) => {
		throw new Error(`samples/manifest.json: ${sample?.id ?? '?'} ${why}`);
	};
	if (!Array.isArray(manifest?.samples)) throw new Error('samples/manifest.json lists no samples');
	const seen = new Set();
	for (const sample of manifest.samples) {
		for (const key of ['id', 'product', 'title', 'description', 'file', 'thumbnail', 'sha256'])
			if (typeof sample?.[key] !== 'string' || !sample[key]) fail(sample, `has no ${key}`);
		if (!SAMPLE_ID.test(sample.id)) fail(sample, 'has an id that is not a lower-case slug');
		if (seen.has(sample.id)) fail(sample, 'is listed twice');
		seen.add(sample.id);
		if (!PRODUCTS.some((product) => product.id === sample.product))
			fail(sample, `names the product ${sample.product}`);
		if (
			!Array.isArray(sample.features) ||
			sample.features.length === 0 ||
			sample.features.some((feature) => typeof feature !== 'string' || !feature)
		)
			fail(sample, 'needs a list of feature strings');
		if (!Number.isSafeInteger(sample.size) || sample.size <= 0)
			fail(sample, 'needs its size in bytes');
		if (!/^[0-9a-f]{64}$/u.test(sample.sha256)) fail(sample, 'has a sha256 that is not one');
		// The suite builds the request from these, so they stay inside samples/.
		for (const key of ['file', 'thumbnail'])
			if (!/^[a-z]+\/[a-z0-9-]+\.[a-z]+$/u.test(sample[key]))
				fail(sample, `has an unexpected ${key} path`);
	}
	return manifest;
}

/** The manifest, checked for the fields the page and the suite read. */
export function readManifest(root = ROOT) {
	return checkManifest(JSON.parse(readFileSync(join(root, 'samples', 'manifest.json'), 'utf8')));
}

function card(sample, product) {
	const name = sample.file.split('/').pop();
	const features = sample.features.map((feature) => `<li>${escapeHtml(feature)}</li>`).join('');
	return `<article class="hub-card sample-card" id="sample-${sample.id}" data-sample="${sample.id}">
								<div class="sample-shot">
									<img
										src="${sample.thumbnail}"
										alt="Preview of ${escapeHtml(sample.title)}"
										width="${THUMBNAIL.width}"
										height="${THUMBNAIL.height}"
										loading="lazy"
										decoding="async"
									/>
								</div>
								<span class="app-badge" style="--app: ${product.colour}">${product.letter}</span>
								<h3>${escapeHtml(sample.title)}</h3>
								<p>${escapeHtml(sample.description)}</p>
								<ul class="sample-features" aria-label="What it shows">${features}</ul>
								<p class="sample-file"><code>${escapeHtml(name)}</code> · ${formatSize(sample.size)}</p>
								<nav class="hub-links" aria-label="${escapeHtml(sample.title)}">
									<a class="sample-open" href="../?sample=${sample.id}">Open in ${product.app}</a
									><a href="${sample.file}" download>Download</a>
								</nav>
							</article>`;
}

/** The gallery page for `manifest`, from the template's three markers. */
export function renderSamplesPage(template, manifest) {
	const shown = PRODUCTS.map((product) => ({
		product,
		samples: manifest.samples.filter((sample) => sample.product === product.id),
	})).filter((group) => group.samples.length > 0);
	const links = shown
		.map(({ product }) => `<a href="#${product.app.toLowerCase()}">${product.app}</a>`)
		.join('\n\t\t\t\t\t');
	const sections = shown
		.map(
			({ product, samples }) => `<section id="${product.app.toLowerCase()}" class="hub-section">
						<div class="hub-section-bar">
							<h2>${product.app}</h2>
							<span>${countWord(samples.length)} ${product.noun[samples.length === 1 ? 0 : 1]}</span>
						</div>
						<div class="hub-grid sample-grid">
							${samples.map((sample) => card(sample, product)).join('\n\t\t\t\t\t\t\t')}
						</div>
					</section>`,
		)
		.join('\n\n\t\t\t\t\t');
	const fill = { nav: links, outline: links, sections, count: countWord(manifest.samples.length) };
	return template.replace(/<!-- samples:(\w+) -->/g, (marker, name) => {
		if (!(name in fill)) throw new Error(`${TEMPLATE}: unknown marker ${marker}`);
		return fill[name];
	});
}

/**
 * Writes the gallery into `<out>/samples/`: the page, its stylesheet, the manifest the suite
 * reads when a sample is opened, every sample file and its thumbnail. Returns the sample ids.
 */
export function writeSamplesPage(out, { root = ROOT } = {}) {
	const manifest = readManifest(root);
	const target = join(out, 'samples');
	mkdirSync(target, { recursive: true });
	// The template is source, not a page (build-pages.mjs copies the whole of site/).
	rmSync(join(target, 'template.html'), { force: true });
	copyFileSync(join(root, 'site', 'samples', 'samples.css'), join(target, 'samples.css'));
	writeFileSync(
		join(target, 'index.html'),
		renderSamplesPage(readFileSync(join(root, TEMPLATE), 'utf8'), manifest),
	);
	copyFileSync(join(root, 'samples', 'manifest.json'), join(target, 'manifest.json'));
	for (const sample of manifest.samples)
		for (const path of [sample.file, sample.thumbnail]) {
			const from = join(root, 'samples', path);
			if (!existsSync(from)) throw new Error(`samples/${path} does not exist`);
			mkdirSync(dirname(join(target, path)), { recursive: true });
			copyFileSync(from, join(target, path));
		}
	return manifest.samples.map((sample) => sample.id);
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
	const out = resolve(ROOT, process.argv[2] ?? 'suite-dist');
	console.log(
		`Samples gallery: ${writeSamplesPage(out).length} samples in ${join(out, 'samples')}`,
	);
}
