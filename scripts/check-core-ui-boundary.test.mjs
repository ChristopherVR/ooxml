import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { checkCoreUiBoundary } from './check-core-ui-boundary.mjs';

const SCRIPT = fileURLToPath(new URL('./check-core-ui-boundary.mjs', import.meta.url));

function fixture(t, files, manifest = {}) {
	const root = mkdtempSync(join(tmpdir(), 'ooxml-core-ui-boundary-'));
	// The target is the unique temporary directory created above.
	t.after(() => rmSync(root, { recursive: true, force: true }));
	mkdirSync(join(root, 'src'));
	writeFileSync(join(root, 'package.json'), JSON.stringify(manifest));
	for (const [name, source] of Object.entries(files))
		writeFileSync(join(root, 'src', name), source);
	return root;
}

for (const source of [
	`import { Button } from 'ooxml-ui';`,
	`import 'ooxml-ui';`,
	`import type { Button } from 'ooxml-ui/components';`,
	`import\n/* boundary */ { Button }\nfrom\n'ooxml-ui';`,
	`export { Button } from 'ooxml-ui';`,
	`export * from 'ooxml-ui/components';`,
	`export type { Button } from 'ooxml-ui';`,
	`const ui = import('ooxml-ui');`,
	'const ui = import(`ooxml-ui/components`);',
	`const ui = require('ooxml-ui');`,
	`const ui = require.resolve('ooxml-ui/components');`,
	`import ui = require('ooxml-ui');`,
	`type UI = import('ooxml-ui').Button;`,
	`/// <reference types="ooxml-ui" />`,
]) {
	test(`rejects actual module dependency: ${source}`, (t) => {
		const root = fixture(t, { 'model.ts': source });
		const result = spawnSync(process.execPath, [SCRIPT, root], { encoding: 'utf8' });
		assert.equal(result.status, 1, result.stderr);
		assert.match(result.stderr, /model\.ts.*ooxml-ui/);
	});
}

test('comments, documentation and ordinary strings pass the CLI', (t) => {
	const root = fixture(t, {
		'model.ts': `/** The ooxml-ui virtualised worksheet viewport. */
// import { Button } from 'ooxml-ui';
/* export * from 'ooxml-ui'; */
const documentation = "import('ooxml-ui')";
const packageName = 'ooxml-ui';
import { model } from 'ooxml-ui-tools';
`,
		'model.test.ts': `import 'ooxml-ui';`,
	});
	const result = spawnSync(process.execPath, [SCRIPT, root], { encoding: 'utf8' });
	assert.equal(result.status, 0, result.stderr);
	assert.match(result.stdout, /boundary passed/);
});

for (const extension of ['ts', 'tsx', 'mts', 'cts', 'js', 'jsx', 'mjs', 'cjs']) {
	test(`checks .${extension} source files`, (t) => {
		const root = fixture(t, { [`model.${extension}`]: `import 'ooxml-ui/components';` });
		assert.equal(checkCoreUiBoundary(root).length, 1);
	});
}

for (const field of [
	'dependencies',
	'devDependencies',
	'peerDependencies',
	'optionalDependencies',
]) {
	test(`rejects ${field} on the core package, including npm aliases`, (t) => {
		const root = fixture(t, {}, { [field]: { 'ooxml-ui': '*', uiAlias: 'npm:ooxml-ui@^1.0.0' } });
		const result = spawnSync(process.execPath, [SCRIPT, root], { encoding: 'utf8' });
		assert.equal(result.status, 1, result.stderr);
		assert.match(result.stderr, new RegExp(`${field}\\.ooxml-ui`));
		assert.match(result.stderr, new RegExp(`${field}\\.uiAlias`));
	});
}

test('reports all dependency sites with source line numbers', (t) => {
	const root = fixture(t, {
		'model.ts': `// harmless ooxml-ui\nimport 'ooxml-ui';\nexport * from 'ooxml-ui/components';`,
	});
	const problems = checkCoreUiBoundary(root);
	assert.equal(problems.length, 2);
	assert.match(problems[0], /model\.ts:2:/);
	assert.match(problems[1], /model\.ts:3:/);
});
