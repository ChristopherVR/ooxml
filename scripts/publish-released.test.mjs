import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative } from 'node:path';
import { after, test } from 'node:test';

import { PACKAGES } from './release-plan.mjs';
import {
	assertEntryPoints,
	groupOf,
	groupsOf,
	missingEntryPoints,
	publishDirOf,
	publishManifest,
	resolveTargets,
	verifyManifest,
} from './publish-released.mjs';

const [key, meta] = Object.entries(PACKAGES)[0];
const version = JSON.parse(
	await (await import('node:fs/promises')).readFile(`${meta.dir}/package.json`, 'utf8'),
).version;

test('a tag resolves to exactly one package, scoped names included', () => {
	const [target] = resolveTargets({ tag: `${meta.npm}@1.2.3` });
	assert.deepEqual(target, { key, npm: meta.npm, dir: meta.dir, version: '1.2.3' });
});

test('a malformed or unknown tag is rejected', () => {
	assert.throws(() => resolveTargets({ tag: 'v1.2.3' }), /Invalid tag|Unknown package/);
	assert.throws(() => resolveTargets({ tag: 'nope@1.0.0' }), /Unknown package/);
	assert.throws(() => resolveTargets({ tag: `${meta.npm}@1.0.0; rm -rf /` }), /Invalid tag/);
});

test('a plan resolves to its released packages in plan order', () => {
	const plan = {
		order: ['b', 'a', 'c'],
		packages: {
			a: { release: true, npm: '@x/a', dir: 'packages/a', version: '1.0.0' },
			b: { release: false, npm: '@x/b', dir: 'packages/b', version: '1.0.0' },
			c: { release: true, npm: '@x/c', dir: 'packages/c', version: '2.0.0' },
		},
	};
	assert.deepEqual(
		resolveTargets({ plan }).map((t) => t.npm),
		['@x/a', '@x/c'],
	);
});

test('publish groups keep a viewer together and every other package alone', () => {
	const targets = [
		{ key: 'core', dir: 'src/core' },
		{ key: 'ui', dir: 'src/ui' },
		{ key: 'docx-react', dir: 'viewers/docx/packages/react' },
		{ key: 'docx-vue', dir: 'viewers/docx/packages/vue' },
		{ key: 'mcp', dir: 'mcp' },
	];
	assert.equal(groupOf(targets[2]), 'viewers/docx');
	assert.deepEqual(groupsOf(targets), ['core', 'ui', 'viewers/docx', 'mcp']);
});

test('the manifest on disk must be the version being published', () => {
	assert.throws(
		() => verifyManifest({ npm: meta.npm, dir: meta.dir, version: '9.9.9' }),
		/on disk/,
	);
	assert.doesNotThrow(() => verifyManifest({ npm: meta.npm, dir: meta.dir, version }));
});

const temp = mkdtempSync(join(tmpdir(), 'publish-released-'));
after(() => rmSync(temp, { recursive: true, force: true }));
const ui = { npm: '@x/ui', dir: 'src/ui', version: '1.0.0' };
const versions = new Map([['@x/core', '0.4.2']]);
const writeUi = (dependencies, extra = {}) => {
	mkdirSync(join(temp, ui.dir), { recursive: true });
	writeFileSync(
		join(temp, ui.dir, 'package.json'),
		JSON.stringify({ name: ui.npm, version: ui.version, dependencies, ...extra }),
	);
};

test('a workspace range on a sibling is published as a caret range on its version', () => {
	writeUi(
		{ '@x/core': 'workspace:*', lit: '^3.0.0' },
		{ devDependencies: { '@x/core': 'workspace:^' } },
	);
	const manifest = publishManifest(ui, versions, temp);
	assert.deepEqual(manifest.dependencies, { '@x/core': '^0.4.2', lit: '^3.0.0' });
	assert.deepEqual(manifest.devDependencies, { '@x/core': '^0.4.2' });
});

test('a `*` range on a sibling (the UI on the core) is published as a caret range', () => {
	writeUi({ '@x/core': '*', lit: '^3.0.0' });
	assert.deepEqual(publishManifest(ui, versions, temp).dependencies, {
		'@x/core': '^0.4.2',
		lit: '^3.0.0',
	});
});

test('a `file:` link on a sibling (the MCP package on the core) is published as a caret range', () => {
	writeUi({ '@x/core': 'file:..' });
	assert.deepEqual(publishManifest(ui, versions, temp).dependencies, { '@x/core': '^0.4.2' });
});

test('a workspace range on anything but a sibling cannot be published', () => {
	writeUi({ '@x/other': 'workspace:*' });
	assert.throws(() => verifyManifest(ui, versions, temp), /cannot be installed/);
	writeUi({ '@x/other': 'file:../other' });
	assert.throws(() => verifyManifest(ui, versions, temp), /cannot be installed/);
});

test('a real sibling range must be satisfied by the sibling version', () => {
	writeUi({ '@x/core': '^0.4.0' });
	assert.doesNotThrow(() => verifyManifest(ui, versions, temp));
	writeUi({ '@x/core': '^0.3.0' });
	assert.throws(() => verifyManifest(ui, versions, temp), /but that package is at 0.4.2/);
});

test('pre-release verification rejects a peer range that excludes the new core major', () => {
	const major = new Map([['@x/core', '1.0.0']]);
	writeUi({}, { peerDependencies: { '@x/core': '>=0.20.0 <1' } });
	assert.throws(() => verifyManifest(ui, major, temp), /but that package is at 1.0.0/);
	writeUi({}, { peerDependencies: { '@x/core': '^1.0.0' } });
	assert.doesNotThrow(() => verifyManifest(ui, major, temp));
});

test('a workspace devDependency on a private package is dropped from the published manifest', () => {
	writeUi({ lit: '^3.0.0' }, { devDependencies: { '@x/private': 'workspace:*' } });
	assert.deepEqual(publishManifest(ui, versions, temp).devDependencies, {});
});

test('the pptx Angular binding is published from its ng-packagr dist, by tag or by plan', () => {
	const angular = 'viewers/pptx/packages/angular';
	const [byTag] = resolveTargets({ tag: 'pptx-angular-viewer@1.2.3' });
	assert.equal(byTag.publishDir, 'dist');
	assert.equal(publishDirOf(byTag), `${angular}/dist`);
	const plan = {
		order: ['pptx-angular', 'pptx-react'],
		packages: {
			'pptx-angular': { release: true, npm: 'pptx-angular-viewer', dir: angular, version: '1.0.0' },
			'pptx-react': {
				release: true,
				npm: 'pptx-react-viewer',
				dir: 'viewers/pptx/packages/react',
				version: '1.0.0',
			},
		},
	};
	assert.deepEqual(resolveTargets({ plan }).map(publishDirOf), [
		`${angular}/dist`,
		'viewers/pptx/packages/react',
	]);
});

test('a publishDir manifest is the one published; the source manifest is the one verified', () => {
	const target = { npm: '@x/ng', dir: 'packages/ng', publishDir: 'dist', version: '1.0.0' };
	const write = (dir, manifest) => {
		mkdirSync(join(temp, dir), { recursive: true });
		writeFileSync(join(temp, dir, 'package.json'), JSON.stringify({ name: '@x/ng', ...manifest }));
	};
	write('packages/ng', { version: '1.0.0', dependencies: { '@x/core': 'workspace:*' } });
	assert.throws(
		() => publishManifest(target, versions, temp),
		/packages\/ng\/dist\/package.json does not exist: build @x\/ng first/,
	);
	write('packages/ng/dist', {
		version: '1.0.0',
		exports: { '.': './fesm2022/ng.mjs' },
		dependencies: { '@x/core': 'workspace:*' },
	});
	const manifest = publishManifest(target, versions, temp);
	assert.deepEqual(manifest.exports, { '.': './fesm2022/ng.mjs' });
	assert.deepEqual(manifest.dependencies, { '@x/core': '^0.4.2' });
	write('packages/ng/dist', { version: '0.9.0' });
	assert.throws(() => publishManifest(target, versions, temp), /0\.9\.0 on disk/);
	assert.doesNotThrow(() => verifyManifest(target, versions, temp));
});

test('an absolute target directory is read as is, not under the repository root', () => {
	// On Windows, `relative()` from a checkout on D: to the C: temp directory is this absolute path.
	const dir = join(temp, 'absolute/ng');
	const target = { npm: '@x/abs', dir, publishDir: 'dist', version: '1.0.0' };
	const write = (sub, manifest) => {
		mkdirSync(join(dir, sub), { recursive: true });
		writeFileSync(join(dir, sub, 'package.json'), JSON.stringify({ name: '@x/abs', ...manifest }));
	};
	write('.', { version: '1.0.0' });
	write('dist', { version: '1.0.0', dependencies: { '@x/core': 'workspace:*' } });
	assert.deepEqual(publishManifest(target, versions).dependencies, { '@x/core': '^0.4.2' });
	assert.doesNotThrow(() => verifyManifest(target, versions));
});

// The manifest and files of pptx-angular-viewer 4.25.0 to 4.61.0: the source directory, unbuilt.
const brokenAngular = {
	name: 'pptx-angular-viewer',
	exports: {
		'./styles': './pptx-angular-viewer.css',
		'./internals': {
			types: './types/pptx-angular-viewer.d.ts',
			default: './fesm2022/pptx-angular-viewer.mjs',
		},
		'./i18n/fr': { import: './i18n/fr.mjs', require: './i18n/fr.cjs' },
	},
};
const brokenFiles = ['package.json', 'README.md', 'ng-package.json', 'scripts/finalize-dist.mjs'];

test('the entry point check rejects a package without its entry files or a "." export', () => {
	assert.deepEqual(missingEntryPoints(brokenAngular, brokenFiles), [
		'exports "."',
		'exports["./styles"] ./pptx-angular-viewer.css',
		'exports["./internals"] ./types/pptx-angular-viewer.d.ts',
		'exports["./internals"] ./fesm2022/pptx-angular-viewer.mjs',
		'exports["./i18n/fr"] ./i18n/fr.mjs',
		'exports["./i18n/fr"] ./i18n/fr.cjs',
	]);
	assert.deepEqual(
		missingEntryPoints(
			{ main: 'dist/index.cjs', module: './dist/index.mjs', types: './dist/index.d.ts' },
			['package.json', 'dist/index.mjs'],
		),
		['main dist/index.cjs', 'types ./dist/index.d.ts'],
	);
});

test('the entry point check accepts every target form a complete package uses', () => {
	const files = [
		'package.json',
		'index.js',
		'index.d.ts',
		'fesm2022/ng.mjs',
		'i18n/fr.mjs',
		'i18n/de.mjs',
		'assets/logo.svg',
	];
	assert.deepEqual(missingEntryPoints({ main: 'index', exports: './index.js' }, files), []);
	assert.deepEqual(missingEntryPoints({ exports: { import: './index.js' } }, files), []);
	assert.deepEqual(
		missingEntryPoints(
			{
				typings: 'index.d.ts',
				exports: {
					'.': { types: './index.d.ts', default: ['./fesm2022/ng.mjs'] },
					'./i18n/*': './i18n/*.mjs',
					'./assets/': './assets/',
					'./internal/*': null,
					'./package.json': './package.json',
				},
			},
			files,
		),
		[],
	);
	assert.deepEqual(
		missingEntryPoints({ exports: { '.': './index.js', './x/*': './x/*.js' } }, files),
		['exports["./x/*"] ./x/*.js'],
	);
});

test('the entry point check reads the files npm would pack', () => {
	const write = (dir, manifest, files = []) => {
		mkdirSync(join(temp, dir), { recursive: true });
		writeFileSync(join(temp, dir, 'package.json'), JSON.stringify(manifest));
		for (const file of files) {
			mkdirSync(join(temp, dir, file, '..'), { recursive: true });
			writeFileSync(join(temp, dir, file), '');
		}
		return join(temp, dir);
	};
	const manifest = {
		name: '@x/packed',
		version: '1.0.0',
		files: ['dist'],
		exports: { '.': { types: './dist/index.d.ts', default: './dist/index.js' } },
	};
	const good = write('packed/good', manifest, ['dist/index.js', 'dist/index.d.ts']);
	assert.doesNotThrow(() => assertEntryPoints(good, manifest));
	const missing = write('packed/missing', manifest, ['dist/index.js']);
	assert.throws(
		() => assertEntryPoints(missing, manifest),
		/@x\/packed@1\.0\.0 packed from .* is missing exports\["."\] \.\/dist\/index\.d\.ts\.$/,
	);
	// On disk but left out by `files`: npm would not ship it.
	const outside = { ...manifest, types: './types/index.d.ts' };
	const unshipped = write('packed/unshipped', outside, [
		'dist/index.js',
		'dist/index.d.ts',
		'types/index.d.ts',
	]);
	assert.throws(
		() => assertEntryPoints(unshipped, outside),
		/missing types \.\/types\/index\.d\.ts/,
	);
});

test('the entry point check covers bin targets and the svelte field', () => {
	const files = ['package.json', 'index.js', 'cli.js'];
	assert.deepEqual(
		missingEntryPoints({ main: 'index.js', bin: './cli.js', svelte: 'index.js' }, files),
		[],
	);
	assert.deepEqual(
		missingEntryPoints(
			{ bin: { tool: 'cli.js', other: './bin/other.js' }, svelte: './src/index.js' },
			files,
		),
		['bin ./bin/other.js', 'svelte ./src/index.js'],
	);
	assert.deepEqual(missingEntryPoints({ bin: 'bin/cli.js' }, files), ['bin bin/cli.js']);
});

test('the entry point check lists a package past the default 1 MiB spawnSync buffer', () => {
	// 4,000 files with long names: about 1.2 MB of `npm pack --dry-run --json` output.
	const dir = join(temp, 'packed/large');
	const names = Array.from({ length: 4000 }, (_, i) => `dist/${'x'.repeat(200)}-${i}.js`);
	mkdirSync(join(dir, 'dist'), { recursive: true });
	for (const name of names) writeFileSync(join(dir, name), '');
	const manifest = { name: '@x/large', version: '1.0.0', exports: { '.': `./${names[0]}` } };
	writeFileSync(join(dir, 'package.json'), JSON.stringify(manifest));
	assert.doesNotThrow(() => assertEntryPoints(dir, manifest));
});

test("a failing npm pack --dry-run reports npm's error", () => {
	const dir = join(temp, 'packed/invalid');
	mkdirSync(dir, { recursive: true });
	writeFileSync(join(dir, 'package.json'), '{ "name": ');
	assert.throws(
		() => assertEntryPoints(dir, { name: '@x/invalid', version: '1.0.0' }),
		/npm pack --dry-run failed in .*invalid: .*EJSONPARSE/su,
	);
});

test('--entry-points checks every built package of a plan before anything is tagged', async () => {
	const { spawnSync } = await import('node:child_process');
	const dir = join(temp, 'plan/pkg');
	mkdirSync(join(dir, 'dist'), { recursive: true });
	writeFileSync(join(dir, 'dist/index.js'), '');
	const manifest = { name: '@x/planned', version: '1.0.0', exports: { '.': './dist/index.js' } };
	const plan = join(temp, 'plan/release-plan.json');
	const write = (packageManifest) => {
		writeFileSync(join(dir, 'package.json'), JSON.stringify(packageManifest));
		writeFileSync(
			plan,
			JSON.stringify({
				order: ['planned'],
				packages: {
					planned: { release: true, npm: '@x/planned', dir: relative('.', dir), version: '1.0.0' },
				},
			}),
		);
	};
	const run = () =>
		spawnSync(
			process.execPath,
			// `--dry-run` too: without `--entry-points` the script would go on to publish.
			['scripts/publish-released.mjs', '--plan', plan, '--entry-points', '--dry-run'],
			{ encoding: 'utf8' },
		);
	write(manifest);
	const passed = run();
	assert.equal(passed.status, 0, passed.stderr);
	assert.match(passed.stdout, /Checked the entry points of 1 built packages\./u);
	write({ ...manifest, types: './dist/index.d.ts' });
	const failed = run();
	assert.notEqual(failed.status, 0);
	assert.match(
		failed.stderr,
		/@x\/planned@1\.0\.0 packed from .* is missing types \.\/dist\/index\.d\.ts/u,
	);
});

test('the unbuilt pptx Angular source directory fails the entry point check', () => {
	const dir = 'viewers/pptx/packages/angular';
	const manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
	assert.throws(
		() => assertEntryPoints(dir, manifest),
		/missing exports ".", .*\.\/fesm2022\/pptx-angular-viewer\.mjs/,
	);
});

test('--from-dir without a packed.json index fails before uploading anything', async () => {
	const { spawnSync } = await import('node:child_process');
	const result = spawnSync(
		process.execPath,
		['scripts/publish-released.mjs', '--tag', `${meta.npm}@${version}`, '--from-dir', temp],
		{ encoding: 'utf8' },
	);
	assert.notEqual(result.status, 0);
});
