import { readFileSync } from 'node:fs';
import { relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { balanceFiles } from './browser-shards.mjs';

// New files can be split from expensive suites. Budget more than a typical small test file and
// spread them across shards until a fresh report provides actual timings.
const UNKNOWN_FILE_SECONDS = 5;

export function balanceTestFiles(files, timings, count) {
	const weights = Object.fromEntries(
		files.map((file) => [file, timings[file] ?? UNKNOWN_FILE_SECONDS]),
	);
	return balanceFiles(files, weights, count);
}

// The package supplies its own Vitest class so the script works with the isolated linker.
export function durationSequencer(BaseSequencer, timingFile) {
	const root = fileURLToPath(new URL('.', timingFile));
	return class extends BaseSequencer {
		async shard(specifications) {
			const { index, count } = this.ctx.config.shard;
			const timings = JSON.parse(readFileSync(timingFile, 'utf8'));
			const paths = specifications.map((spec) =>
				relative(root, spec.moduleId).split(sep).join('/'),
			);
			const shards = balanceTestFiles(paths, timings, count);
			const selected = new Set(shards[index - 1].files);
			return specifications.filter((_, i) => selected.has(paths[i]));
		}
	};
}
