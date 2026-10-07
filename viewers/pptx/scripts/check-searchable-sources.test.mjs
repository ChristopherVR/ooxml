import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { findUnsearchableSources } from './check-searchable-sources.mjs';

const NUL = String.fromCharCode(0);

function fixture(files) {
	const root = mkdtempSync(join(tmpdir(), 'searchable-'));
	for (const [relativePath, contents] of Object.entries(files)) {
		const full = join(root, relativePath);
		mkdirSync(join(full, '..'), { recursive: true });
		writeFileSync(full, contents, 'utf8');
	}
	return root;
}

test('reports a source file carrying a raw NUL byte', () => {
	const root = fixture({
		'viewers/pptx/packages/a/src/keys.ts': `export const SEP = '${NUL}';\n`,
	});
	const offences = findUnsearchableSources(root);
	assert.deepEqual(offences, [{ file: 'viewers/pptx/packages/a/src/keys.ts', line: 1 }]);
});

test('reports the line the byte sits on, once per occurrence', () => {
	const root = fixture({
		'e2e/pptx/support/keys.ts': `const a = 1;\nconst b = '${NUL}';\nconst c = '${NUL}';\n`,
	});
	assert.deepEqual(findUnsearchableSources(root), [
		{ file: 'e2e/pptx/support/keys.ts', line: 2 },
		{ file: 'e2e/pptx/support/keys.ts', line: 3 },
	]);
});

test('accepts the escaped spelling, which is the same string at runtime', () => {
	const root = fixture({
		'viewers/pptx/packages/a/src/keys.ts': "export const SEP = '\\u0000';\n",
	});
	assert.deepEqual(findUnsearchableSources(root), []);
});

test('checks the migrated renderer and document operations', () => {
	const root = fixture({
		'src/ui/src/pptx/render/keys.ts': `const key = '${NUL}';\n`,
		'src/core/pptx/editor/render/keys.ts': `const key = '${NUL}';\n`,
	});
	assert.deepEqual(findUnsearchableSources(root), [
		{ file: 'src/ui/src/pptx/render/keys.ts', line: 1 },
		{ file: 'src/core/pptx/editor/render/keys.ts', line: 1 },
	]);
});

test('ignores build output and vendored copies', () => {
	const root = fixture({
		'viewers/pptx/packages/a/dist/index.js': `'${NUL}'`,
		'viewers/pptx/packages/a/node_modules/dep/index.js': `'${NUL}'`,
		'viewers/pptx/packages/angular/src/internal/shared-src/keys.ts': `'${NUL}'`,
	});
	assert.deepEqual(findUnsearchableSources(root), []);
});

test('the repository itself is searchable', () => {
	assert.deepEqual(findUnsearchableSources(), []);
});
