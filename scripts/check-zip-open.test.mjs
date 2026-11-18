import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { findRawZipOpens } from './check-zip-open.mjs';

const SCRIPT = fileURLToPath(new URL('./check-zip-open.mjs', import.meta.url));
const RAW = 'export const open = (bytes) => JSZip.loadAsync(bytes);\n';

function fixture(t, files) {
	const root = mkdtempSync(join(tmpdir(), 'ooxml-zip-open-'));
	// The target is the unique temporary directory created above.
	t.after(() => rmSync(root, { recursive: true, force: true }));
	for (const [name, source] of Object.entries(files)) {
		mkdirSync(dirname(join(root, name)), { recursive: true });
		writeFileSync(join(root, name), source);
	}
	return root;
}

test('reports a raw open in shipped source of every root', (t) => {
	const root = fixture(t, {
		'src/core/pptx/load.ts': `\n${RAW}`,
		'src/ui/src/open.ts': RAW,
		'apps/office-suite/src/open.js': RAW,
		'mcp/src/open.js': RAW,
		'viewers/docx/packages/react/src/open.tsx': RAW,
	});
	assert.deepEqual(findRawZipOpens(root).sort(), [
		'apps/office-suite/src/open.js:1',
		'mcp/src/open.js:1',
		'src/core/pptx/load.ts:2',
		'src/ui/src/open.ts:1',
		'viewers/docx/packages/react/src/open.tsx:1',
	]);
});

test('allows the guard itself, tests, fixtures and scripts', (t) => {
	const root = fixture(t, {
		'src/core/opc/zip-limits.ts': RAW,
		'src/core/opc/a.test.ts': RAW,
		'src/core/docx/test-support/inspect.ts': RAW,
		'src/core/pptx/__tests__/helpers/deck.ts': RAW,
		'src/core/pptx/core/deck.test-util.ts': RAW,
		'src/core/xlsx/write/mini-package-fixtures.ts': RAW,
		'src/core/scripts/build.ts': RAW,
		'viewers/pptx/scripts/make.mjs': RAW,
		'src/core/node_modules/jszip/index.js': RAW,
	});
	assert.deepEqual(findRawZipOpens(root), []);
});

test('the command fails on a raw open and names it', (t) => {
	const root = fixture(t, { 'src/core/load.ts': RAW });
	const result = spawnSync(process.execPath, [SCRIPT, root], { encoding: 'utf8' });
	assert.equal(result.status, 1);
	assert.match(result.stderr, /src\/core\/load\.ts:1/);
});

test('this repository has no raw open', () => {
	assert.deepEqual(findRawZipOpens(), []);
});
