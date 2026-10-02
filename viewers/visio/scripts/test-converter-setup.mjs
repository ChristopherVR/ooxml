import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import {
	copyFileSync,
	existsSync,
	mkdirSync,
	mkdtempSync,
	readFileSync,
	rmSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const source = resolve(dirname(fileURLToPath(import.meta.url)), 'setup-converter.mjs');
function fixture(t) {
	const dir = mkdtempSync(join(tmpdir(), 'visio-converter-setup-'));
	t.after(() => rmSync(dir, { recursive: true, force: true }));
	const viewer = join(dir, 'viewer'),
		core = join(dir, 'emf-converter-current');
	for (const path of [join(viewer, 'scripts'), join(viewer, 'integration'), core, join(dir, 'bin')])
		mkdirSync(path, { recursive: true });
	copyFileSync(source, join(viewer, 'scripts/setup-converter.mjs'));
	copyFileSync(
		resolve(dirname(source), 'converter-snapshot.mjs'),
		join(viewer, 'scripts/converter-snapshot.mjs'),
	);
	const git = (...args) =>
		execFileSync('git', args, { cwd: core, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
	git('init');
	git('config', 'user.name', 'Fixture');
	git('config', 'user.email', 'fixture@example.invalid');
	writeFileSync(join(core, 'source.txt'), 'baseline\n');
	git('add', 'source.txt');
	git('commit', '-m', 'fixture');
	writeFileSync(join(viewer, 'integration/emf-converter-revision.txt'), git('rev-parse', 'HEAD'));
	writeFileSync(join(core, 'source.txt'), 'patched\n');
	writeFileSync(join(viewer, 'integration/emf-converter-visio.patch'), git('diff', '--binary'));
	writeFileSync(join(viewer, 'integration/emf-converter-patched-revision.txt'), '1'.repeat(40));
	const lock = '{"name":"fixture","lockfileVersion":3}\n';
	writeFileSync(join(viewer, 'integration/emf-converter-package-lock.json'), lock);
	const calls = join(dir, 'npm-calls');
	writeFileSync(join(dir, 'bin/npm'), '#!/bin/sh\nprintf "%s\\n" "$*" >> "$SETUP_TEST_CALLS"\n', {
		mode: 0o755,
	});
	const run = () =>
		spawnSync(process.execPath, [join(viewer, 'scripts/setup-converter.mjs')], {
			encoding: 'utf8',
			env: {
				...process.env,
				PATH: `${join(dir, 'bin')}:${process.env.PATH}`,
				SETUP_TEST_CALLS: calls,
			},
		});
	return { core, viewer, git, run, calls, lock };
}

test('verified checkout installs pinned lock then builds, and is repeatable', (t) => {
	const f = fixture(t);
	for (let i = 0; i < 2; i++) {
		const result = f.run();
		assert.equal(result.status, 0, result.stderr);
		assert.equal(readFileSync(join(f.core, 'package-lock.json'), 'utf8'), f.lock);
	}
	assert.deepEqual(readFileSync(f.calls, 'utf8').trim().split('\n'), [
		'ci --ignore-scripts',
		'run build',
		'ci --ignore-scripts',
		'run build',
	]);
});
for (const [name, mutate, message] of [
	[
		'wrong revision',
		(f) => writeFileSync(join(f.viewer, 'integration/emf-converter-revision.txt'), '0'.repeat(40)),
		/different revision/,
	],
	[
		'invalid pin',
		(f) => writeFileSync(join(f.viewer, 'integration/emf-converter-revision.txt'), 'main'),
		/Invalid pinned/,
	],
	['staged changes', (f) => f.git('add', 'source.txt'), /staged changes/],
	[
		'modified patch',
		(f) => writeFileSync(join(f.core, 'source.txt'), 'user change\n'),
		/do not match/,
	],
	[
		'untracked source',
		(f) => writeFileSync(join(f.core, 'extra.txt'), 'user file\n'),
		/do not match/,
	],
	[
		'different lock',
		(f) => writeFileSync(join(f.core, 'package-lock.json'), '{"different":true}\n'),
		/lock differs/,
	],
]) {
	test(`refuses ${name} without running npm or overwriting files`, (t) => {
		const f = fixture(t);
		mutate(f);
		const before = f.git('diff', '--binary');
		const result = f.run();
		assert.notEqual(result.status, 0);
		assert.match(result.stderr, message);
		assert.equal(f.git('diff', '--binary'), before);
		assert.equal(existsSync(f.calls), false);
	});
}

test('equivalent lock formatting is accepted without overwriting it', (t) => {
	const f = fixture(t);
	const formatted = JSON.stringify(JSON.parse(f.lock), null, 2);
	writeFileSync(join(f.core, 'package-lock.json'), formatted);
	const result = f.run();
	assert.equal(result.status, 0, result.stderr);
	assert.equal(readFileSync(join(f.core, 'package-lock.json'), 'utf8'), formatted);
});

test('accepts the exact committed patch revision but rejects later edits', (t) => {
	const f = fixture(t);
	f.git('add', 'source.txt');
	f.git('commit', '-m', 'patch');
	writeFileSync(
		join(f.viewer, 'integration/emf-converter-patched-revision.txt'),
		f.git('rev-parse', 'HEAD'),
	);
	assert.equal(f.run().status, 0);
	writeFileSync(join(f.core, 'source.txt'), 'later edit\n');
	const result = f.run();
	assert.notEqual(result.status, 0);
	assert.match(result.stderr, /do not match/);
});

test('accepts fresh patch additions regardless of tracked/untracked ordering', (t) => {
	const f = fixture(t);
	writeFileSync(join(f.core, 'a-added.txt'), 'added first alphabetically\n');
	f.git('add', 'source.txt', 'a-added.txt');
	const base = readFileSync(
		join(f.viewer, 'integration/emf-converter-revision.txt'),
		'utf8',
	).trim();
	writeFileSync(
		join(f.viewer, 'integration/emf-converter-visio.patch'),
		f.git('diff', '--binary', base),
	);
	f.git('reset');
	assert.equal(f.run().status, 0);
});
