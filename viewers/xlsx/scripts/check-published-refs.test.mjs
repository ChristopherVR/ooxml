import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
	forbiddenManifestEntries,
	importedPackages,
	undeclaredImports,
} from './check-published-refs.mjs';

const manifest = {
	dependencies: { '@christophervr/xlsx-core': '^0.1.0', 'ooxml-core': '^0.3.0' },
	peerDependencies: { react: '>=18' },
};

test('finds static, dynamic and re-export specifiers but skips relative ones', () => {
	const source = [
		"import a from 'react';",
		'export * from "ooxml-ui/controls";',
		"const b = await import('@christophervr/ole2/ole2-parser-read');",
		"import './chunk.js';",
		'const msg = \'x from": "y\';',
	].join('\n');
	assert.deepEqual(importedPackages(source), [
		'react',
		'ooxml-ui/controls',
		'@christophervr/ole2/ole2-parser-read',
	]);
});

test('declared dependencies, subpaths and node built-ins are allowed', () => {
	const source = [
		"import { x } from '@christophervr/xlsx-core';",
		"import { y } from '@christophervr/xlsx-core/load';",
		"import 'ooxml-core/xlsx';",
		"import fs from 'node:fs';",
		"import 'react';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), []);
});

test('flags inlined internal packages and ole2 that leaked into a tarball', () => {
	const source = [
		"import 'jszip';",
		"export * from 'xlsx-web-component';",
		"import 'xlsx-bindings/react';",
		"import '@christophervr/ole2/legacy-excel-workbook';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), [
		'jszip',
		'xlsx-web-component',
		'xlsx-bindings/react',
		'@christophervr/ole2/legacy-excel-workbook',
	]);
});

test('rejects internal, ole2 and workspace entries in a published manifest', () => {
	assert.deepEqual(forbiddenManifestEntries(manifest), []);
	// ooxml-ui is a registry dependency like @christophervr/xlsx-core: never inlined, never forbidden.
	assert.deepEqual(forbiddenManifestEntries({ dependencies: { 'ooxml-ui': '^0.1.1' } }), []);
	assert.deepEqual(
		forbiddenManifestEntries({
			dependencies: {
				'@christophervr/ole2': '0.2.0',
				'xlsx-bindings': '^0.1.0',
				'xlsx-vue-viewer': '^0.1.0',
				'@christophervr/xlsx-react-viewer': '^0.1.0',
				'@christophervr/xlsx-core': 'workspace:*',
				'ooxml-core': 'file:../../../ooxml-xlsx',
				xlsx: '^0.18.5',
			},
		}),
		[
			'dependencies.@christophervr/ole2',
			'dependencies.xlsx-bindings',
			'dependencies.xlsx-vue-viewer',
			'dependencies.@christophervr/xlsx-react-viewer',
			'dependencies.@christophervr/xlsx-core',
			'dependencies.ooxml-core',
		],
	);
});
