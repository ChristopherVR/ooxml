import { readdirSync, readFileSync } from 'node:fs';
import { join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { BaseSequencer } from 'vitest/node';
import { durationSequencer } from '../../scripts/test-sequencer.mjs';
import { testWorkers } from '../../scripts/test-workers.mjs';

const root = fileURLToPath(new URL('.', import.meta.url));

// Test files share a worker (`isolate: false`): evaluating the module graph once per file was most
// of the suite's time. A file that mocks a module, or asserts the absence of browser globals (SSR),
// depends on a fresh module registry, so it keeps a worker of its own in `core-isolated`. They are
// found by content, so a new `vi.mock` test is isolated without touching this file.
function testFiles(dir: string, out: string[] = []): string[] {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		if (entry.name === 'node_modules' || entry.name === 'dist' || entry.name.startsWith('.'))
			continue;
		const path = join(dir, entry.name);
		if (entry.isDirectory()) testFiles(path, out);
		else if (entry.name.endsWith('.test.ts')) out.push(path);
	}
	return out;
}
const rel = (file: string) => relative(root, file).split(sep).join('/');
const needsIsolation: string[] = [];
// A module cached by one file extends that file's `HTMLElement`, so a file pinned to another
// DOM (`// @vitest-environment jsdom`) must not share a worker with the node default: each
// environment is its own project, with its own workers and module cache.
const byEnvironment = new Map<string, string[]>();
for (const file of testFiles(root)) {
	const source = readFileSync(file, 'utf8');
	if (file.endsWith('ssr.test.ts') || /\bvi\.(?:mock|doMock)\s*[<(]/.test(source)) {
		needsIsolation.push(rel(file));
		continue;
	}
	const environment = /@vitest-environment\s+([\w-]+)/.exec(source)?.[1] ?? 'node';
	if (environment !== 'node') {
		byEnvironment.set(environment, [...(byEnvironment.get(environment) ?? []), rel(file)]);
	}
}
const otherEnvironments = [...byEnvironment.values()].flat();

// The pptx tests read real decks from pptx/__tests__/fixtures (including the e2e snapshot that is
// committed under fixtures/e2e), so nothing outside this repository is needed.
const common = {
	resolve: {
		alias: [
			{
				find: /^#xpath-runtime$/,
				replacement: fileURLToPath(new URL('./xml/xpath-runtime.ts', import.meta.url)),
			},
			{
				find: /^#inkml-runtime$/,
				replacement: fileURLToPath(new URL('./drawingml/inkml/runtime.ts', import.meta.url)),
			},
			{
				find: /^#embedded-workbook-chart$/,
				replacement: fileURLToPath(
					new URL('./xlsx/ui/dialogs/embedded-chart-entry.ts', import.meta.url),
				),
			},
			{
				find: /^ooxml-core\/math$/,
				replacement: fileURLToPath(new URL('./math/index.ts', import.meta.url)),
			},
			{
				find: /^ooxml-core\/pptx$/,
				replacement: fileURLToPath(new URL('./pptx/index.ts', import.meta.url)),
			},
		],
	},
};
const project = (
	name: string,
	include: string[],
	exclude: string[],
	isolate: boolean,
	environment = 'node',
) => ({
	...common,
	test: {
		name,
		isolate,
		globals: true,
		environment,
		include,
		exclude: ['node_modules', 'dist', ...exclude],
		// The pptx area's integration tests load real multi-megabyte decks through the full
		// parse pipeline (and several round-trip them through save); the 5s default timed out on CI.
		testTimeout: 30_000,
		hookTimeout: 30_000,
	},
});

export default {
	test: {
		maxWorkers: testWorkers(),
		sequence: {
			sequencer: durationSequencer(
				BaseSequencer,
				new URL('./vitest.timings.json', import.meta.url),
			),
		},
		projects: [
			project('core', ['**/*.test.ts'], [...needsIsolation, ...otherEnvironments], false),
			...[...byEnvironment].map(([environment, files]) =>
				project(`core-${environment}`, files, [], false, environment),
			),
			...(needsIsolation.length > 0 ? [project('core-isolated', needsIsolation, [], true)] : []),
		],
	},
};
