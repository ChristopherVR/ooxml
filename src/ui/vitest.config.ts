import { readdirSync, readFileSync } from 'node:fs';
import { sep } from 'node:path';
import { fileURLToPath } from 'node:url';

// Development resolves the core package from the repository root source, so UI tests never need
// a build of the core. Published consumers get the real dependency.
const coreSource = fileURLToPath(new URL('../core', import.meta.url)).replaceAll(sep, '/');
const uiSource = fileURLToPath(new URL('./src', import.meta.url)).replaceAll(sep, '/');

// Test files share a worker (`isolate: false`): evaluating the module graph once per file was most
// of the suite's time. A file that mocks a module, or asserts the absence of browser globals (SSR),
// depends on a fresh module registry, so it keeps a worker of its own in the `*-isolated` projects.
// They are found by content, so a new `vi.mock` test is isolated without touching this file.
const needsIsolation = readdirSync(fileURLToPath(new URL('./src', import.meta.url)), {
	recursive: true,
})
	.map((file) => `src/${String(file).split(sep).join('/')}`)
	.filter(
		(file) =>
			file.endsWith('.test.ts') &&
			(file.endsWith('ssr.test.ts') ||
				/\bvi\.(?:mock|doMock)\(/.test(
					readFileSync(new URL(`./${file}`, import.meta.url), 'utf8'),
				)),
	);
const isolated = new Set(needsIsolation);
const only = (pptx: boolean) =>
	needsIsolation.filter((file) => file.startsWith('src/pptx/') === pptx);

// Component styles are `.css?raw` text; return the real CSS so tests can read it.
const css = { include: [/\.css\?raw$/] };

// happy-dom is the default DOM (about twice as fast to set up); the files that need jsdom behaviour
// (shadow-root focus, global confirm, and similar) say so with // @vitest-environment jsdom.
const ui = {
	resolve: {
		alias: [
			{ find: /^ooxml-core\/(.+)$/, replacement: `${coreSource}/$1/index.ts` },
			// The package imports its own subpaths by name; resolve them to source, not to dist.
			{ find: /^ooxml-ui\/pptx\/dom$/, replacement: `${uiSource}/pptx/dom/index.ts` },
		],
	},
	test: {
		globals: true,
		// Editor-mounting tests build a full ProseMirror view; under a parallel run they can take 5s+.
		testTimeout: 20_000,
		css,
		environment: 'happy-dom',
		setupFiles: ['./vitest.setup.ts'],
	},
};
const pptx = {
	resolve: { alias: [{ find: /^ooxml-core\/(.+)$/, replacement: `${coreSource}/$1` }] },
	test: {
		css,
		globals: true,
		environment: 'node',
		testTimeout: 30_000,
		hookTimeout: 30_000,
	},
};
const dom = (name: string, include: string[], exclude: string[], isolate: boolean) => ({
	...ui,
	test: { ...ui.test, name, isolate, include, exclude: ['node_modules', 'dist', ...exclude] },
});
const node = (name: string, include: string[], exclude: string[], isolate: boolean) => ({
	...pptx,
	test: { ...pptx.test, name, isolate, include, exclude: ['node_modules', 'dist', ...exclude] },
});

// All four projects share one worker pool, so `vitest run` schedules the DOM tests and the pptx
// tests together. `--project ui*` or `--project pptx*` runs one family.
export default {
	test: {
		// Uncapped, every core runs a worker and the heaviest imports (manifest, SSR) time out; 8 is
		// both stable and the fastest measured.
		maxWorkers: 8,
		projects: [
			dom('ui', ['src/**/*.test.ts'], ['src/pptx/**', ...isolated], false),
			dom('ui-isolated', only(false), [], true),
			node('pptx', ['src/pptx/**/*.test.ts'], [...isolated], false),
			node('pptx-isolated', only(true), [], true),
		],
	},
};
