/**
 * The publishable packages of the viewers imported under `viewers/<name>`, as `release-plan.mjs`
 * table entries keyed `<viewer>-<key>`.
 *
 * Each viewer builds its packages with its own `scripts/build-packages.mjs`, which inlines the
 * viewer's private directories (`bundled`) into every framework package. A change to a bundled
 * directory therefore releases every package marked `bundled`, and a change to one of the
 * viewer's shared build files (`globals`) releases every package of that viewer, except one
 * marked `global: false` (the teams server ships its own source).
 *
 * The pptx bindings also inline the core code they import (their bundles have no `ooxml-core`
 * import), so `inlinedCore` lists those core paths, relative to the repository root, as further
 * triggers of every bundled package. Before the viewer moved here, every core release reached
 * them through a dependency bump instead.
 */

const FRAMEWORKS = ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];

/** `packages/<framework>` bindings published as `name(framework)`, all bundling the viewer's internals. */
const frameworks = (name) =>
	Object.fromEntries(
		FRAMEWORKS.map((framework) => [
			framework,
			{ dir: `packages/${framework}`, npm: name(framework), bundled: true },
		]),
	);

/**
 * The entries of one viewer. `bundled` and `globals` are paths relative to `viewers/<viewer>`;
 * `inlinedCore` is relative to the repository root.
 */
function viewer(name, { bundled, inlinedCore = [], globals, packages }) {
	const at = (path) => `viewers/${name}/${path}`;
	return Object.fromEntries(
		Object.entries(packages).map(([key, meta]) => [
			`${name}-${key}`,
			{
				dir: at(meta.dir),
				npm: meta.npm,
				...(meta.bundled ? { triggers: [...bundled.map(at), ...inlinedCore] } : {}),
				...(meta.global === false ? {} : { globals: globals.map(at) }),
			},
		]),
	);
}

/**
 * The core areas (and the pptx bundler configs) the pptx bindings inline: the `pptx` area, the
 * areas it imports, and the subpaths the viewer imports directly.
 */
const PPTX_INLINED_CORE = [
	...['pptx', 'automation', 'chart', 'color', 'crypto', 'diagram', 'geometry', 'math', 'opc'].map(
		(area) => `src/core/${area}`,
	),
	...['text', 'units', 'xml'].map((area) => `src/core/${area}`),
	'src/core/tsup.pptx.config.ts',
	'src/core/tsup.pptx-editor.config.ts',
	'src/core/tsdown.pptx.config.ts',
];

/** pptx publishes five bindings (no Solid) that inline core/locales and import public UI. */
const PPTX_FRAMEWORKS = ['react', 'vue', 'angular', 'svelte', 'vanilla'];

export const VIEWER_PACKAGES = {
	...viewer('docx', {
		bundled: ['packages/web-component', 'packages/bindings'],
		globals: ['scripts/build-packages.mjs', 'tsconfig.release.json'],
		packages: {
			mcp: { dir: 'mcp', npm: 'docx-viewer-mcp' },
			core: { dir: 'packages/core', npm: 'docx-core' },
			...frameworks((framework) => `docx-${framework}-viewer`),
		},
	}),
	...viewer('xlsx', {
		bundled: ['packages/web-component', 'packages/bindings'],
		globals: ['scripts/build-packages.mjs', 'tsconfig.release.json'],
		packages: {
			mcp: { dir: 'mcp', npm: 'xlsx-viewer-mcp' },
			core: { dir: 'packages/core', npm: '@christophervr/xlsx-core' },
			...frameworks((framework) =>
				framework === 'react' ? '@christophervr/xlsx-react-viewer' : `xlsx-${framework}-viewer`,
			),
		},
	}),
	...viewer('visio', {
		bundled: ['src', 'packages/bindings/src'],
		globals: ['scripts/build-packages.mjs', 'tsconfig.build.json'],
		packages: {
			mcp: { dir: 'mcp', npm: 'visio-viewer-mcp' },
			core: { dir: 'packages/core', npm: 'visio-core' },
			...frameworks((framework) => `visio-${framework}-viewer`),
		},
	}),
	...viewer('pptx', {
		bundled: ['packages/core', 'packages/shared', 'packages/locales'],
		inlinedCore: PPTX_INLINED_CORE,
		globals: ['tsconfig.json', 'scripts/bundle-declarations.mjs'],
		packages: {
			core: { dir: 'packages/core', npm: 'pptx-viewer-core' },
			mcp: { dir: 'packages/tools', npm: 'pptx-viewer-mcp' },
			cli: { dir: 'packages/cli', npm: '@christophervr/pptx-viewer' },
			...Object.fromEntries(
				PPTX_FRAMEWORKS.map((framework) => [
					framework,
					{ dir: `packages/${framework}`, npm: `pptx-${framework}-viewer`, bundled: true },
				]),
			),
		},
	}),
	...viewer('teams', {
		bundled: ['packages/web-component'],
		globals: ['scripts/build-packages.mjs', 'tsconfig.release.json'],
		packages: {
			...frameworks((framework) => `openteams-${framework}-viewer`),
			server: { dir: 'server', npm: 'openteams-server', global: false },
		},
	}),
};

/** Directory (under the repository root) of the viewer a table entry belongs to, or null. */
export const viewerOf = (dir) => /^viewers\/[^/]+/u.exec(dir)?.[0] ?? null;
