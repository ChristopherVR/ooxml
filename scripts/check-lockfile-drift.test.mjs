import assert from 'node:assert/strict';
import { test } from 'node:test';

import { lockDrift } from './check-lockfile-drift.mjs';

const lock = (version, extra = '') => `{
  "lockfileVersion": 1,
  "workspaces": {
    "src/core": {
      "name": "ooxml-core",
      "version": "${version}",
      "dependencies": {
        "yjs": "^13.6.0",
      },
    },
  },
  "packages": {
    "yjs": ["yjs@13.6.0", "", {}, "sha512-aaa"],${extra}
  }
}`;

test('a workspace version bump alone is not drift', () => {
	assert.deepEqual(lockDrift(lock('1.0.0'), lock('1.0.1')), []);
});

test('a changed third-party resolution is drift', () => {
	const moved = lock('1.0.1')
		.replace('yjs@13.6.0', 'yjs@13.6.9')
		.replace('sha512-aaa', 'sha512-bbb');
	assert.deepEqual(lockDrift(lock('1.0.0'), moved), ['bun.lock.packages.yjs']);
});

test('a new package is drift', () => {
	const added = lock('1.0.1', '\n    "evil": ["evil@1.0.0", "", {}, "sha512-ccc"],');
	assert.deepEqual(lockDrift(lock('1.0.0'), added), ['bun.lock.packages.evil']);
});

test('a changed workspace dependency range is drift', () => {
	const ranged = lock('1.0.1').replace('^13.6.0', '^13.7.0');
	assert.deepEqual(lockDrift(lock('1.0.0'), ranged), [
		'bun.lock.workspaces.src/core.dependencies.yjs',
	]);
});

/** The lock with a viewer workspace whose dependencies are `deps`. */
const withViewer = (deps) =>
	lock('1.0.1').replace(
		'"workspaces": {',
		`"workspaces": {
    "viewer": { "name": "pptx-react-viewer", "dependencies": ${JSON.stringify(deps)} },`,
	);

test('a raised range on another workspace package is not drift', () => {
	const viewer = (range) => withViewer({ 'ooxml-core': range });
	assert.deepEqual(lockDrift(viewer('^1.0.0'), viewer('^1.0.1')), []);
});

test('a raised third-party range next to a workspace one is still drift', () => {
	const viewer = (range) => withViewer({ 'ooxml-core': range, yjs: range });
	assert.deepEqual(lockDrift(viewer('^1.0.0'), viewer('^1.0.1')), [
		'bun.lock.workspaces.viewer.dependencies.yjs',
	]);
});
