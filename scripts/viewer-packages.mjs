/**
 * The publishable packages of the viewers imported under `viewers/<name>`, as `release-plan.mjs`
 * table entries keyed `<viewer>-<key>`.
 *
 * Each viewer builds its packages with its own `scripts/build-packages.mjs`, which inlines the
 * viewer's private directories (`bundled`) into every framework package. A change to a bundled
 * directory therefore releases every package marked `bundled`, and a change to one of the
 * viewer's shared build files (`globals`) releases every package of that viewer, except one
 * marked `global: false` (the teams server ships its own source).
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

/** The entries of one viewer. `bundled` and `globals` are paths relative to `viewers/<viewer>`. */
function viewer(name, { bundled, globals, packages }) {
	const at = (path) => `viewers/${name}/${path}`;
	return Object.fromEntries(
		Object.entries(packages).map(([key, meta]) => [
			`${name}-${key}`,
			{
				dir: at(meta.dir),
				npm: meta.npm,
				...(meta.bundled ? { triggers: bundled.map(at) } : {}),
				...(meta.global === false ? {} : { globals: globals.map(at) }),
			},
		]),
	);
}

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
