// Packs the built package, installs the tarball into a clean project and imports every entry
// point the way a consumer would (ESM `import`, and CJS `require` where the entry declares one).
// Run after `bun run build`. Uses `npm pack`/`npm install` only, so it behaves the same on every OS.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const manifest = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
const work = await mkdtemp(path.join(tmpdir(), 'ooxml-core-smoke-'));
const npmCli =
	process.platform === 'win32'
		? path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
		: undefined;
const npm = 'npm';
const run = (command, args, cwd) => {
	const executable = command === 'npm' && npmCli ? process.execPath : command;
	const executableArgs = command === 'npm' && npmCli ? [npmCli, ...args] : args;
	const result = spawnSync(executable, executableArgs, { cwd, encoding: 'utf8' });
	if (result.status !== 0)
		throw new Error(`${command} ${args.join(' ')} failed\n${result.stdout}\n${result.stderr}`);
	return result.stdout;
};

try {
	const packed = JSON.parse(
		run(npm, ['pack', '--json', '--ignore-scripts', '--pack-destination', work], root),
	)[0];
	const files = new Set(packed.files.map((entry) => entry.path));
	for (const required of ['LICENSE', 'NOTICE', 'README.md', 'dist/index.js', 'dist/index.d.ts'])
		assert(files.has(required), `the package is missing ${required}`);
	assert(
		![...files].some((file) => /__tests__|\.test\.|fixtures\//.test(file)),
		'tests or fixtures were packed',
	);
	assert(
		!JSON.stringify(manifest.dependencies).includes('file:'),
		'a runtime dependency uses file:',
	);

	await writeFile(
		path.join(work, 'package.json'),
		JSON.stringify({ private: true, type: 'module' }),
	);
	run(
		npm,
		[
			'install',
			'--ignore-scripts',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			// The optional peers some entry points import (signatures, canvas rasterisation).
			...Object.entries(manifest.peerDependencies ?? {}).map(([name, range]) => `${name}@${range}`),
			path.join(work, packed.filename),
		],
		work,
	);

	const subpaths = Object.keys(manifest.exports).filter((entry) => !entry.endsWith('/cli'));
	const checks = subpaths
		.map((entry) => {
			const specifier = path.posix.join(manifest.name, entry === '.' ? '' : entry);
			const declaresRequire =
				typeof manifest.exports[entry] === 'object' && 'require' in manifest.exports[entry];
			return [
				`assert.ok(Object.keys(await import('${specifier}')).length > 0, '${specifier} has no exports');`,
				declaresRequire
					? `assert.ok(Object.keys(require('${specifier}')).length > 0, '${specifier} (cjs) has no exports');`
					: '',
			].join('\n');
		})
		.join('\n');
	await mkdir(path.join(work, 'check'), { recursive: true });
	await writeFile(
		path.join(work, 'consumer.mjs'),
		`import assert from 'node:assert/strict';\nimport { createRequire } from 'node:module';\nconst require = createRequire(import.meta.url);\n${checks}\nconsole.log('all entry points import');\n`,
	);
	console.log(run('node', ['consumer.mjs'], work).trim());
} finally {
	await rm(work, { recursive: true, force: true });
}
