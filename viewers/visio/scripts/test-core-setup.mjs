import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import {
	copyFileSync,
	existsSync,
	mkdirSync,
	mkdtempSync,
	readFileSync,
	readlinkSync,
	rmSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const scripts = dirname(fileURLToPath(import.meta.url));
function fixture(t) {
	const dir = mkdtempSync(join(tmpdir(), 'visio-core-setup-'));
	t.after(() => rmSync(dir, { recursive: true, force: true }));
	const viewer = join(dir, 'viewer'),
		core = join(dir, 'ooxml');
	for (const path of [join(viewer, 'scripts'), join(viewer, 'integration'), core, join(dir, 'bin')])
		mkdirSync(path, { recursive: true });
	for (const file of [
		'setup-core.mjs',
		'core-paths.mjs',
		'core-snapshot.mjs',
		'link-core.mjs',
		'update-core-patch.mjs',
	])
		copyFileSync(join(scripts, file), join(viewer, 'scripts', file));
	const git = (...args) =>
		execFileSync('git', args, { cwd: core, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
	git('init');
	git('config', 'user.name', 'Fixture');
	git('config', 'user.email', 'fixture@example.invalid');
	writeFileSync(join(core, 'source.txt'), 'baseline\n');
	writeFileSync(
		join(core, 'package.json'),
		JSON.stringify({ name: 'ooxml-core', exports: { './visio': './dist/visio/index.mjs' } }),
	);
	git('add', '.');
	git('commit', '-m', 'fixture');
	writeFileSync(join(viewer, 'integration/core-revision.txt'), git('rev-parse', 'HEAD'));
	const lock = '{"name":"fixture","lockfileVersion":3}\n';
	writeFileSync(join(viewer, 'integration/core-package-lock.json'), lock);
	const calls = join(dir, 'npm-calls');
	writeFileSync(join(dir, 'bin/npm'), '#!/bin/sh\nprintf "%s\\n" "$*" >> "$SETUP_TEST_CALLS"\n', {
		mode: 0o755,
	});
	const run = (env = {}, script = 'setup-core.mjs') => {
		const inherited = { ...process.env };
		delete inherited.VISIO_CORE_DIR;
		delete inherited.VISIO_NPM_CACHE;
		return spawnSync(process.execPath, [join(viewer, 'scripts', script)], {
			encoding: 'utf8',
			env: {
				...inherited,
				PATH: `${join(dir, 'bin')}:${process.env.PATH}`,
				SETUP_TEST_CALLS: calls,
				...env,
			},
		});
	};
	return { dir, core, viewer, git, run, calls, lock };
}

test('clean published checkout needs no patch and installs/builds repeatably', (t) => {
	const f = fixture(t);
	for (let i = 0; i < 2; i++) {
		const result = f.run();
		assert.equal(result.status, 0, result.stderr);
		assert.equal(readFileSync(join(f.core, 'package-lock.json'), 'utf8'), f.lock);
	}
	assert.deepEqual(readFileSync(f.calls, 'utf8').trim().split('\n'), [
		'ci --ignore-scripts',
		'run build:visio',
		'ci --ignore-scripts',
		'run build:visio',
	]);
});
for (const [name, mutate, message] of [
	[
		'wrong revision',
		(f) => writeFileSync(join(f.viewer, 'integration/core-revision.txt'), '0'.repeat(40)),
		/different revision/,
	],
	[
		'invalid pin',
		(f) => writeFileSync(join(f.viewer, 'integration/core-revision.txt'), 'main'),
		/Invalid pinned/,
	],
	[
		'staged changes',
		(f) => {
			writeFileSync(join(f.core, 'source.txt'), 'staged\n');
			f.git('add', 'source.txt');
		},
		/staged changes/,
	],
	[
		'modified source',
		(f) => writeFileSync(join(f.core, 'source.txt'), 'user change\n'),
		/do not match/,
	],
	[
		'untracked source',
		(f) => writeFileSync(join(f.core, 'extra file.txt'), 'user file\n'),
		/do not match/,
	],
	[
		'different lock',
		(f) => writeFileSync(join(f.core, 'package-lock.json'), '{"different":true}\n'),
		/lock differs/,
	],
])
	test(`refuses ${name} without npm or overwriting files`, (t) => {
		const f = fixture(t);
		mutate(f);
		const before = f.git('diff', '--binary', 'HEAD');
		const lock = existsSync(join(f.core, 'package-lock.json'))
			? readFileSync(join(f.core, 'package-lock.json'), 'utf8')
			: undefined;
		const result = f.run();
		assert.notEqual(result.status, 0);
		assert.match(result.stderr, message);
		assert.equal(f.git('diff', '--binary', 'HEAD'), before);
		assert.equal(existsSync(f.calls), false);
		assert.equal(
			existsSync(join(f.core, 'package-lock.json'))
				? readFileSync(join(f.core, 'package-lock.json'), 'utf8')
				: undefined,
			lock,
		);
	});

test('optional patch round-trips tracked and untracked additions independent of ordering', (t) => {
	const f = fixture(t);
	writeFileSync(join(f.core, 'source.txt'), 'patched\n');
	writeFileSync(join(f.core, 'a-added.txt'), 'new\n');
	const patchResult = f.run({}, 'update-core-patch.mjs');
	assert.equal(patchResult.status, 0, patchResult.stderr);
	const result = f.run();
	assert.equal(result.status, 0, result.stderr);
	f.git('restore', 'source.txt');
	rmSync(join(f.core, 'a-added.txt'));
	assert.equal(f.run({}, 'update-core-patch.mjs').status, 0);
	assert.equal(existsSync(join(f.viewer, 'integration/ooxml-visio.patch')), false);
});

test('accepts equivalent npm lock JSON without overwriting formatting', (t) => {
	const f = fixture(t),
		formatted = JSON.stringify(JSON.parse(f.lock), null, 2);
	writeFileSync(join(f.core, 'package-lock.json'), formatted);
	const result = f.run();
	assert.equal(result.status, 0, result.stderr);
	assert.equal(readFileSync(join(f.core, 'package-lock.json'), 'utf8'), formatted);
});

test('override clones pinned source separately without changing canonical sibling', (t) => {
	const f = fixture(t),
		selected = join(f.dir, 'isolated');
	const result = f.run({
		VISIO_CORE_DIR: '../isolated',
		GIT_CONFIG_COUNT: '1',
		GIT_CONFIG_KEY_0: `url.${f.core}.insteadOf`,
		GIT_CONFIG_VALUE_0: 'https://github.com/ChristopherVR/ooxml.git',
	});
	assert.equal(result.status, 0, result.stderr);
	assert.equal(readFileSync(join(selected, 'package-lock.json'), 'utf8'), f.lock);
	assert.equal(existsSync(join(f.core, 'package-lock.json')), false);
	assert.equal(f.git('status', '--porcelain'), '');
});

test('local link is explicit, repeatable and refuses a real installed directory', (t) => {
	const f = fixture(t),
		link = join(f.viewer, 'node_modules/ooxml-core');
	for (let i = 0; i < 2; i++) {
		const result = f.run({ VISIO_CORE_DIR: f.core }, 'link-core.mjs');
		assert.equal(result.status, 0, result.stderr);
	}
	assert.equal(resolve(dirname(link), readlinkSync(link)), f.core);
	rmSync(link);
	mkdirSync(link);
	writeFileSync(join(link, 'user.txt'), 'preserve');
	const result = f.run({}, 'link-core.mjs');
	assert.notEqual(result.status, 0);
	assert.match(result.stderr, /not a symlink/);
	assert.equal(readFileSync(join(link, 'user.txt'), 'utf8'), 'preserve');
});
