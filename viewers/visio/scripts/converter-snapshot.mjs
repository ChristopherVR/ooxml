import { execFileSync, spawnSync } from 'node:child_process';

// Git emits staged/tracked additions together, but git apply leaves new paths
// untracked. Compare the same patch sections regardless of checkout representation.
export function canonicalPatch(patch) {
	return patch
		.replace(/\r\n/g, '\n')
		.split(/(?=^diff --git )/m)
		.filter(Boolean)
		.sort()
		.join('');
}
export function converterSnapshot(core, revision) {
	const git = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
	let patch = git(['diff', '--binary', revision]);
	for (const file of git(['ls-files', '--others', '--exclude-standard', '-z'])
		.split('\0')
		.filter(Boolean)) {
		if (file === 'package-lock.json') continue;
		const diff = spawnSync('git', ['diff', '--no-index', '--binary', '--', '/dev/null', file], {
			cwd: core,
			encoding: 'utf8',
		});
		if (diff.status !== 0 && diff.status !== 1)
			throw new Error('Cannot verify untracked converter file.');
		patch += diff.stdout;
	}
	return canonicalPatch(patch);
}
