import { readdirSync, readFileSync } from 'node:fs';
import { join, relative, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

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
const needsIsolation = testFiles(root)
	.filter(
		(file) =>
			file.endsWith('ssr.test.ts') ||
			/\bvi\.(?:mock|doMock)\s*[<(]/.test(readFileSync(file, 'utf8')),
	)
	.map((file) => relative(root, file).split(sep).join('/'));

// The pptx tests read real decks from pptx/__tests__/fixtures (including the e2e snapshot that is
// committed under fixtures/e2e), so nothing outside this repository is needed.
const common = {
	resolve: {
		alias: [
			{
				find: /^ooxml-core\/pptx$/,
				replacement: fileURLToPath(new URL('./pptx/index.ts', import.meta.url)),
			},
		],
	},
};
const project = (name: string, include: string[], exclude: string[], isolate: boolean) => ({
	...common,
	test: {
		name,
		isolate,
		globals: true,
		environment: 'node',
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
		maxWorkers: 4,
		projects: [
			project('core', ['**/*.test.ts'], needsIsolation, false),
			...(needsIsolation.length > 0 ? [project('core-isolated', needsIsolation, [], true)] : []),
		],
	},
};
