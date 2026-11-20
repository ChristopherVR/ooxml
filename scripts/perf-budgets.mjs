#!/usr/bin/env node
/**
 * Byte budgets CI enforces: what each published entry point and each suite screen costs to
 * download, gzipped, against the limits in `scripts/perf-budgets.json`.
 *
 *   node scripts/perf-budgets.mjs check     measure and fail on a budget broken, stale or missing
 *   node scripts/perf-budgets.mjs update    rewrite every limit from a fresh measurement
 *   node scripts/perf-budgets.mjs measure   print the measurement as JSON (--files: and what
 *                                           each screen requested)
 *
 * Needs `src/core/dist` and `src/ui/dist` (`bun run build`, `bun run --cwd src/ui build`) and
 * Playwright's Chromium; the suite is built here into a temporary directory. Every number is a
 * byte count of build output, so it is the same on any machine with the same lockfile. See
 * docs/performance-budgets.md.
 */
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { measurePackage } from './bundle-size.mjs';
import { measureSuiteSite, scriptFiles } from './perf-budget-routes.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const BUDGETS_FILE = join(ROOT, 'scripts', 'perf-budgets.json');
/** A limit is the measurement plus this share, rounded up to a whole KiB. */
export const HEADROOM = 0.03;
/** A limit this far above what `update` would write hides a regression of that size: it fails. */
export const STALE = 0.1;
const KIB = 1024;

/**
 * Refuses output that names the checkout it was built in (a bundler comment with an absolute
 * path, say): its size would then depend on the machine, and it would publish that path.
 */
export function assertNoCheckoutPath(name, code, root = ROOT) {
	for (const spelling of new Set([root, root.replaceAll('\\', '/')]))
		if (code.includes(spelling)) throw new Error(`${name} contains the checkout path ${spelling}.`);
}

/** Every measurement: the published entry points of ooxml-core and ooxml-ui, and the suite. */
export async function measure({ onRoute } = {}) {
	const sizes = {};
	for (const dir of ['src/core', 'src/ui']) {
		for (const file of scriptFiles(join(ROOT, dir, 'dist')))
			assertNoCheckoutPath(relative(ROOT, file), readFileSync(file, 'utf8'));
		for (const [name, entry] of Object.entries(measurePackage(join(ROOT, dir))))
			sizes[`package/${name}`] = entry.gzip;
	}
	const { buildSuite } = await import(pathToFileURL(join(ROOT, 'scripts/build-suite.mjs')).href);
	const site = mkdtempSync(join(tmpdir(), 'ooxml-perf-budgets-'));
	try {
		await buildSuite(site, { logLevel: 'warn' });
		for (const file of scriptFiles(site))
			assertNoCheckoutPath(relative(site, file), readFileSync(file, 'utf8'));
		Object.assign(sizes, await measureSuiteSite(site, { onRoute }));
	} finally {
		rmSync(site, { recursive: true, force: true });
	}
	return sizes;
}

/** The limit `update` writes for a measurement: 3% over it, at least 1 KiB, in whole KiB. */
export const limitFor = (bytes) =>
	Math.max(Math.ceil((bytes * (1 + HEADROOM)) / KIB), Math.ceil(bytes / KIB) + 1) * KIB;

/**
 * What is wrong with a measurement against the limits: a budget broken, a limit so far above the
 * measurement that a regression would hide under it, a measurement with no budget, a budget with
 * nothing measured.
 */
export function problems(sizes, limits) {
	const out = [];
	const kib = (bytes) => `${(bytes / KIB).toFixed(1)} KiB`;
	for (const [name, bytes] of Object.entries(sizes)) {
		const limit = limits[name];
		if (limit === undefined) out.push(`${name}: ${kib(bytes)} has no budget.`);
		else if (bytes > limit)
			out.push(
				`${name}: ${kib(bytes)} is over its budget of ${kib(limit)} by ${kib(bytes - limit)}.`,
			);
		else if (limit > limitFor(bytes) * (1 + STALE))
			out.push(
				`${name}: ${kib(bytes)} is far under its budget of ${kib(limit)}; lower it to keep the win.`,
			);
	}
	for (const name of Object.keys(limits))
		if (!(name in sizes)) out.push(`${name}: has a budget but was not measured.`);
	return out;
}

export function readLimits(file = BUDGETS_FILE) {
	return JSON.parse(readFileSync(file, 'utf8')).budgets;
}

export function writeLimits(sizes, file = BUDGETS_FILE) {
	const budgets = Object.fromEntries(
		Object.entries(sizes)
			.sort(([a], [b]) => a.localeCompare(b))
			.map(([name, bytes]) => [name, limitFor(bytes)]),
	);
	const $comment =
		'Gzipped byte limits. Written by `node scripts/perf-budgets.mjs update`; read docs/performance-budgets.md before changing one.';
	writeFileSync(file, `${JSON.stringify({ $comment, budgets }, null, '\t')}\n`);
}

async function main(command, flags) {
	if (command === 'measure') {
		const files = {};
		const sizes = await measure({ onRoute: (name, list) => (files[name] = list) });
		console.log(JSON.stringify(flags.includes('--files') ? { sizes, files } : sizes, null, '\t'));
		return 0;
	}
	if (command === 'update') {
		writeLimits(await measure());
		console.log(`Wrote ${relative(process.cwd(), BUDGETS_FILE)}. Say why in the commit.`);
		return 0;
	}
	if (command === 'check') {
		const sizes = await measure();
		const limits = readLimits();
		for (const [name, bytes] of Object.entries(sizes))
			console.log(`${name.padEnd(48)} ${String(bytes).padStart(10)} / ${limits[name] ?? '-'}`);
		const found = problems(sizes, limits);
		if (!found.length) {
			console.log(`All ${Object.keys(sizes).length} byte budgets hold.`);
			return 0;
		}
		for (const line of found) console.error(process.env.GITHUB_ACTIONS ? `::error::${line}` : line);
		console.error(
			'Fix a regression in the code. To change a budget on purpose, run `node scripts/perf-budgets.mjs update` and say why in the commit (docs/performance-budgets.md).',
		);
		return 1;
	}
	console.error('Usage: node scripts/perf-budgets.mjs check | update | measure [--files]');
	return 2;
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href)
	process.exitCode = await main(process.argv[2], process.argv.slice(3));
