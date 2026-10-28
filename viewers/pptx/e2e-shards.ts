import { readFileSync, readdirSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { balanceFiles, fileDurations } from '../../scripts/browser-shards.mjs';

function collectSpecs(dir: string, prefix = ''): string[] {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		// Do not descend into the isolated linker's dependency tree, even on runtimes that follow
		// directory symlinks during recursive readdir. Playwright excludes these directories too.
		if (
			['node_modules', 'fixtures', '.git', 'test-results', 'playwright-report'].includes(entry.name)
		)
			return [];
		const file = `${prefix}${entry.name}`;
		if (entry.isDirectory()) return collectSpecs(join(dir, entry.name), `${file}/`);
		return /\.(spec|test)\.[cm]?[jt]sx?$/.test(file) &&
			!/(^|\/)capture-/.test(file) &&
			file !== 'packaged-present.spec.ts'
			? [file]
			: [];
	});
}

let snapshot: { files: string[]; history: unknown } | undefined;
function inputs() {
	if (snapshot) return snapshot;
	let history: unknown;
	try {
		history = JSON.parse(
			readFileSync(resolve(__dirname, '../../.browser-timings/history.json'), 'utf8'),
		);
	} catch {
		/* No history on the first run: deterministically balance files. */
	}
	snapshot = { files: collectSpecs(resolve(__dirname, '../../e2e/pptx')), history };
	return snapshot;
}

/** Filter before importing specs, so every CI shard collects only the files it executes. */
export function shardMatch(project: string): string[] | undefined {
	const shard = process.env.PPTX_E2E_SHARD;
	if (!shard) return undefined;
	if (process.argv.some((arg) => /^--shard(?:=|$)/.test(arg)))
		throw new Error('Use PPTX_E2E_SHARD or --shard, not both');
	const match = /^(\d+)\/(\d+)$/.exec(shard);
	const index = Number(match?.[1]);
	const total = Number(match?.[2]);
	if (!index || index > total || total < 1) throw new Error(`Invalid PPTX_E2E_SHARD: ${shard}`);
	const selected = process.env.PPTX_E2E_SPECS?.split(' ').filter(Boolean);
	const { files: allFiles, history } = inputs();
	const files = allFiles.filter((file) => !selected?.length || selected.includes(file));
	if (selected?.some((file) => !files.includes(file)))
		throw new Error('PPTX_E2E_SPECS names a missing spec');
	const binding = project.replace(/-headless$/, '');
	const totals = fileDurations(history, binding);
	const headless = fileDurations(history, `${binding}-headless`);
	for (const [file, duration] of Object.entries(headless))
		totals[file] = (totals[file] ?? 0) + duration;
	const assigned = balanceFiles(files, totals, total)[index - 1]!.files;
	const matches = project.endsWith('-headless')
		? assigned.filter((file) => file === 'host-owned-collaboration.spec.ts')
		: assigned;
	return matches.length ? matches.map((file) => `**/${file}`) : ['**/__empty-browser-shard__'];
}
