import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import {
	existsSync,
	mkdirSync,
	mkdtempSync,
	readFileSync,
	rmSync,
	symlinkSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

test('standalone MCP copies preserve sources and tests without workspace dependency links', () => {
	const root = mkdtempSync(join(tmpdir(), 'mcp-copy-'));
	try {
		const source = join(root, 'source');
		const target = join(root, 'consumer');
		mkdirSync(join(source, 'src'), { recursive: true });
		mkdirSync(join(source, 'test'));
		mkdirSync(join(source, 'node_modules', '@scope'), { recursive: true });
		writeFileSync(join(source, 'package.json'), '{"type":"module"}');
		writeFileSync(join(source, 'src', 'index.js'), 'export const value = 1;');
		writeFileSync(join(source, 'test', 'index.test.js'), 'source test');
		symlinkSync(
			join(root, 'missing-dependency'),
			join(source, 'node_modules', '@scope', 'sdk'),
			'junction',
		);
		execFileSync(process.execPath, [
			fileURLToPath(new URL('./copy-mcp-sources.mjs', import.meta.url)),
			source,
			target,
		]);
		assert.equal(readFileSync(join(target, 'package.json'), 'utf8'), '{"type":"module"}');
		assert.equal(readFileSync(join(target, 'src', 'index.js'), 'utf8'), 'export const value = 1;');
		assert.equal(readFileSync(join(target, 'test', 'index.test.js'), 'utf8'), 'source test');
		assert.equal(existsSync(join(target, 'node_modules')), false);
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
});

test('a workspace sibling is installed from a tarball of the local package', () => {
	const root = mkdtempSync(join(tmpdir(), 'mcp-copy-'));
	try {
		const source = join(root, 'source');
		const target = join(root, 'consumer');
		mkdirSync(source);
		writeFileSync(
			join(source, 'package.json'),
			JSON.stringify({ dependencies: { 'ooxml-mcp': 'workspace:*', zod: '^4.0.0' } }),
		);
		execFileSync(process.execPath, [
			fileURLToPath(new URL('./copy-mcp-sources.mjs', import.meta.url)),
			source,
			target,
		]);
		const { dependencies } = JSON.parse(readFileSync(join(target, 'package.json'), 'utf8'));
		assert.match(dependencies['ooxml-mcp'], /^file:\.\/\.workspace\/ooxml-mcp-[\d.]+\.tgz$/);
		assert.equal(dependencies.zod, '^4.0.0');
		const tarball = dependencies['ooxml-mcp'].slice('file:./'.length);
		assert.equal(existsSync(join(target, tarball)), true);
	} finally {
		rmSync(root, { recursive: true, force: true });
	}
});
