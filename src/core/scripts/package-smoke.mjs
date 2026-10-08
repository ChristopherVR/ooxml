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
	for (const required of ['LICENSE', 'NOTICE', 'README.md', 'dist/index.mjs', 'dist/index.d.ts'])
		assert(files.has(required), `the package is missing ${required}`);
	assert(
		![...files].some((file) => /__tests__|\.test\.|fixtures\/|test-support\//.test(file)),
		'tests or fixtures were packed',
	);
	assert(
		!JSON.stringify(manifest.dependencies).includes('file:'),
		'a runtime dependency uses file:',
	);

	// The legacy .doc and .xls loaders and the package encryption (its CFB container) inline
	// ole2: consumers must not need it.
	for (const file of [
		'dist/crypto/index.mjs',
		'dist/crypto/index.cjs',
		'dist/docx/load/index.mjs',
		'dist/docx/load/index.cjs',
		'dist/xlsx/load/index.mjs',
		'dist/xlsx/load/index.cjs',
	]) {
		assert(files.has(file), `the package is missing ${file}`);
		const text = await readFile(path.join(root, file), 'utf8');
		assert(!/(?:from|require\()\s*['"]@christophervr\/ole2/.test(text), `${file} imports ole2`);
	}
	assert(!('@christophervr/ole2' in (manifest.dependencies ?? {})), 'ole2 must stay inlined');

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

	const subpaths = Object.keys(manifest.exports)
		.filter((entry) => !entry.endsWith('/cli'))
		.flatMap((entry) => {
			if (!entry.includes('*')) return [entry];
			const target = manifest.exports[entry].import.slice(2);
			const [prefix, suffix] = target.split('*');
			return [...files]
				.filter((file) => file.startsWith(prefix) && file.endsWith(suffix))
				.map((file) => entry.replaceAll('*', file.slice(prefix.length, -suffix.length)));
		});
	const checks = subpaths
		.map((entry) => {
			const specifier = path.posix.join(manifest.name, entry === '.' ? '' : entry);
			const exported =
				manifest.exports[entry] ??
				Object.entries(manifest.exports).find(
					([key]) => key.includes('*') && entry.startsWith(key.split('*')[0]),
				)?.[1];
			const declaresRequire = typeof exported === 'object' && 'require' in exported;
			return [
				`assert.ok(Object.keys(await import('${specifier}')).length > 0 || ${!manifest.exports[entry]}, '${specifier} has no exports');`,
				declaresRequire
					? `assert.ok(Object.keys(require('${specifier}')).length > 0 || ${!manifest.exports[entry]}, '${specifier} (cjs) has no exports');`
					: '',
			].join('\n');
		})
		.join('\n');
	await mkdir(path.join(work, 'check'), { recursive: true });
	await writeFile(
		path.join(work, 'consumer.mjs'),
		`import assert from 'node:assert/strict';\nimport { createRequire } from 'node:module';\nconst require = createRequire(import.meta.url);\n${checks}\n
const Y = await import('yjs');
for (const load of [(s) => import(s), async (s) => require(s)]) {
 const sync = await load('${manifest.name}/pptx/editor/render/collaboration-sync');
 const leases = await load('${manifest.name}/pptx/editor/render/collaboration-text-lease');
 const reconcile = await load('${manifest.name}/pptx/editor/render/collaboration-reconcile');
 const doc = new Y.Doc();
 const factories = { createMap: () => new Y.Map(), createArray: () => new Y.Array(), createText: () => new Y.Text() };
 const element = { id: 'text', type: 'text', x: 0, y: 0, width: 100, height: 20, text: 'before', textSegments: [{ text: 'before', style: {} }] };
 const slides = [{ id: 'slide', slideNumber: 1, elements: [element] }];
 sync.writeSlidesToYDoc(slides, doc, factories);
 const target = doc.getArray(sync.YDOC_SLIDES_KEY).get(0).get('elements').get(0);
 let checked = 0;
 const release = leases.registerCollaborationTextLease(target, () => { checked++; return false; });
 reconcile.reconcileSlidesInYDoc([{ ...slides[0], elements: [{ ...element, text: 'after', textSegments: [{ text: 'after', style: {} }] }] }], doc, factories);
 assert.equal(checked, 1, 'public editor entries must share collaboration leases');
 release(); doc.destroy();
}
console.log('all entry points import');\n`,
	);
	console.log(run('node', ['consumer.mjs'], work).trim());
} finally {
	await rm(work, { recursive: true, force: true });
}
